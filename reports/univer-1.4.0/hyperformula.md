# hyperformula vs univer@1.4.0 (hyperformula)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-lacks-function 241, volatile 10. Fails where the engine's error came from another cell: 3. Fails where both answered an error, but a different one: 45. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (22): ARRAYFORMULA, ARRAY_CONSTRAIN, COUNTUNIQUE, HF.ADD, HF.CONCAT, HF.DIVIDE, HF.EQ, HF.GT, HF.GTE, HF.LT, HF.LTE, HF.MINUS, HF.MULTIPLY, HF.NE, HF.POW, HF.UMINUS, HF.UNARY_PERCENT, HF.UPLUS, INTERVAL, ISBINARY, MAXPOOL, MEDIANPOOL. Never-supported functions present: ARRAY_CONSTRAIN.
Upstream literal differs from Excel's value in 677 of 4491 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 4166 | 3588 | 334 | 3 | 0 | 241 | 91.4% |
| other | 576 | 502 | 62 | 2 | 0 | 10 | 88.7% |
| **total** | 4742 | 4090 | 396 | 5 | 0 | 251 | 91.1% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 4742 | 4090 | 396 | 5 | 0 | 251 | 91.1% |
| **total** | 4742 | 4090 | 396 | 5 | 0 | 251 | 91.1% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| ADDRESS | 62 | fail 62 | `functions/function-address.xlsx!A5` =ADDRESS(1,1,4, TRUE(), "") → Excel "!A1", Univer "A1" |
| TEXT | 18 | fail 18 | `functions/function-text.xlsx!A1` =TEXT(TRUE(), "mm/dd/yyyy") → Excel "TRUE", Univer true |
| HF.EQ | 17 | skip 17 | `functions/function-hfeq.xlsx!A1` =HF.EQ(1) → Excel "#NAME?", Univer "#NAME?" |
| HF.GT | 17 | skip 17 | `functions/function-hfgt.xlsx!A1` =HF.GT(1) → Excel "#NAME?", Univer "#NAME?" |
| HF.GTE | 17 | skip 17 | `functions/function-hfgte.xlsx!A1` =HF.GTE(1) → Excel "#NAME?", Univer "#NAME?" |
| HF.LT | 17 | skip 17 | `functions/function-hflt.xlsx!A1` =HF.LT(1) → Excel "#NAME?", Univer "#NAME?" |
| HF.LTE | 17 | skip 17 | `functions/function-hflte.xlsx!A1` =HF.LTE(1) → Excel "#NAME?", Univer "#NAME?" |
| HF.NE | 17 | skip 17 | `functions/function-hfne.xlsx!A1` =HF.NE(1) → Excel "#NAME?", Univer "#NAME?" |
| INTERVAL | 13 | skip 13 | `functions/function-interval.xlsx!B1` =INTERVAL("12/30/2018") → Excel "#NAME?", Univer "#NAME?" |
| HF.DIVIDE | 12 | skip 12 | `functions/function-hfdivide.xlsx!A1` =HF.DIVIDE(1) → Excel "#NAME?", Univer "#NAME?" |
| COUNTUNIQUE | 11 | skip 11 | `functions/function-countunique.xlsx!A1` =COUNTUNIQUE() → Excel "#NAME?", Univer "#NAME?" |
| HF.ADD | 11 | skip 11 | `functions/function-hfadd.xlsx!A1` =HF.ADD(1) → Excel "#NAME?", Univer "#NAME?" |
| HF.MINUS | 11 | skip 11 | `functions/function-hfminus.xlsx!A1` =HF.MINUS(1) → Excel "#NAME?", Univer "#NAME?" |
| HF.MULTIPLY | 11 | skip 11 | `functions/function-hfmultiply.xlsx!A1` =HF.MULTIPLY(1) → Excel "#NAME?", Univer "#NAME?" |
| HF.UMINUS | 11 | skip 11 | `functions/function-hfuminus.xlsx!A1` =HF.UMINUS() → Excel "#NAME?", Univer "#NAME?" |
| XLOOKUP | 11 | fail 11 | `functions/function-xlookup.xlsx!A4` =XLOOKUP(1, B1:B2, C1:C2, 0, "string") → Excel "#VALUE!", Univer "#REF!" |
| HF.POW | 10 | skip 10 | `functions/function-hfpow.xlsx!A1` =HF.POW(1) → Excel "#NAME?", Univer "#NAME?" |
| HF.UPLUS | 10 | skip 10 | `functions/function-hfuplus.xlsx!A1` =HF.UPLUS() → Excel "#NAME?", Univer "#NAME?" |
| HF.CONCAT | 9 | skip 9 | `functions/function-hfconcat.xlsx!A1` =HF.CONCAT(1) → Excel "#NAME?", Univer "#NAME?" |
| HF.UNARY_PERCENT | 9 | skip 9 | `functions/function-hfunary_percent.xlsx!A1` =HF.UNARY_PERCENT() → Excel "#NAME?", Univer "#NAME?" |
| MATCH | 9 | fail 9 | `functions/function-match.xlsx!A1` =MATCH(C2:C3, B1) → Excel "#VALUE!", Univer "#N/A" |
| SEARCH | 9 | fail 9 | `functions/function-search.xlsx!A5` =SEARCH("?b", "foobarbaz", 5) → Excel 6, Univer "#VALUE!" |
| ARABIC | 7 | fail 7 | `functions/function-arabic.xlsx!A1` =ARABIC(" XD ") → Excel 490, Univer "#VALUE!" |
| SUMPRODUCT | 7 | fail 7 | `functions/function-sumproduct.xlsx!A3` =SUMPRODUCT(A1:B1, TRANSPOSE(A1:A2)) → Excel "#VALUE!", Univer 7 |
| COMPLEX | 6 | fail 6 | `functions/function-complex.xlsx!A11` =COMPLEX(-1, 0) → Excel "-1", Univer -1 |
