# hyperformula vs Apiary 19.0.0 (e018d83eb8c6a829)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-dropped-formula 1, excel-lacks-function 241, volatile 10. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 0. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (22): ARRAYFORMULA, ARRAY_CONSTRAIN, COUNTUNIQUE, HF.ADD, HF.CONCAT, HF.DIVIDE, HF.EQ, HF.GT, HF.GTE, HF.LT, HF.LTE, HF.MINUS, HF.MULTIPLY, HF.NE, HF.POW, HF.UMINUS, HF.UNARY_PERCENT, HF.UPLUS, INTERVAL, ISBINARY, MAXPOOL, MEDIANPOOL. Never-supported functions present: ARRAY_CONSTRAIN.
Upstream literal differs from Excel's value in 677 of 4491 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 4166 | 3920 | 5 | 0 | 0 | 241 | 99.9% |
| other | 577 | 566 | 0 | 0 | 0 | 11 | 100.0% |
| **total** | 4743 | 4486 | 5 | 0 | 0 | 252 | 99.9% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 4743 | 4486 | 5 | 0 | 0 | 252 | 99.9% |
| **total** | 4743 | 4486 | 5 | 0 | 0 | 252 | 99.9% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| HF.EQ | 17 | skip 17 | `functions/function-hfeq.xlsx!A1` =HF.EQ(1) → Excel "#NAME?", GRID "#NAME?" |
| HF.GT | 17 | skip 17 | `functions/function-hfgt.xlsx!A1` =HF.GT(1) → Excel "#NAME?", GRID "#NAME?" |
| HF.GTE | 17 | skip 17 | `functions/function-hfgte.xlsx!A1` =HF.GTE(1) → Excel "#NAME?", GRID "#NAME?" |
| HF.LT | 17 | skip 17 | `functions/function-hflt.xlsx!A1` =HF.LT(1) → Excel "#NAME?", GRID "#NAME?" |
| HF.LTE | 17 | skip 17 | `functions/function-hflte.xlsx!A1` =HF.LTE(1) → Excel "#NAME?", GRID "#NAME?" |
| HF.NE | 17 | skip 17 | `functions/function-hfne.xlsx!A1` =HF.NE(1) → Excel "#NAME?", GRID "#NAME?" |
| INTERVAL | 13 | skip 13 | `functions/function-interval.xlsx!B1` =INTERVAL("12/30/2018") → Excel "#NAME?", GRID "#NAME?" |
| HF.DIVIDE | 12 | skip 12 | `functions/function-hfdivide.xlsx!A1` =HF.DIVIDE(1) → Excel "#NAME?", GRID "#NAME?" |
| COUNTUNIQUE | 11 | skip 11 | `functions/function-countunique.xlsx!A1` =COUNTUNIQUE() → Excel "#NAME?", GRID "#NAME?" |
| HF.ADD | 11 | skip 11 | `functions/function-hfadd.xlsx!A1` =HF.ADD(1) → Excel "#NAME?", GRID "#NAME?" |
| HF.MINUS | 11 | skip 11 | `functions/function-hfminus.xlsx!A1` =HF.MINUS(1) → Excel "#NAME?", GRID "#NAME?" |
| HF.MULTIPLY | 11 | skip 11 | `functions/function-hfmultiply.xlsx!A1` =HF.MULTIPLY(1) → Excel "#NAME?", GRID "#NAME?" |
| HF.UMINUS | 11 | skip 11 | `functions/function-hfuminus.xlsx!A1` =HF.UMINUS() → Excel "#NAME?", GRID "#NAME?" |
| HF.POW | 10 | skip 10 | `functions/function-hfpow.xlsx!A1` =HF.POW(1) → Excel "#NAME?", GRID "#NAME?" |
| HF.UPLUS | 10 | skip 10 | `functions/function-hfuplus.xlsx!A1` =HF.UPLUS() → Excel "#NAME?", GRID "#NAME?" |
| HF.CONCAT | 9 | skip 9 | `functions/function-hfconcat.xlsx!A1` =HF.CONCAT(1) → Excel "#NAME?", GRID "#NAME?" |
| HF.UNARY_PERCENT | 9 | skip 9 | `functions/function-hfunary_percent.xlsx!A1` =HF.UNARY_PERCENT() → Excel "#NAME?", GRID "#NAME?" |
| ISBINARY | 6 | skip 6 | `functions/function-isbinary.xlsx!A1` =ISBINARY("1010") → Excel "#NAME?", GRID "#NAME?" |
| MAXPOOL | 5 | skip 5 | `other/matrix-plugin.xlsx!A4` =maxpool(A1:F3,3) → Excel "#NAME?", GRID "#NAME?" |
| MEDIANPOOL | 5 | skip 5 | `other/matrix-plugin.xlsx!A3` =medianpool(A1:F2,2) → Excel "#NAME?", GRID "#NAME?" |
| ARRAY_CONSTRAIN | 3 | skip 3 | `functions/function-array_constrain.xlsx!A1` =ARRAY_CONSTRAIN(1, 0, 1) → Excel "#NAME?", GRID "#NAME?" |
| TEXT | 3 | fail 3 | `functions/function-text.xlsx!B4` =TEXT(A4, "hh:mm a/p") → Excel "12:30 a", GRID "00:30 a/p" |
| ARRAYFORMULA | 2 | skip 2 | `functions/function-arrayformula.xlsx!A1` =ARRAYFORMULA() → Excel "#NAME?", GRID "#N/A" |
| YEAR | 2 | skip 2 | `functions/function-now.xlsx!A1` =YEAR(NOW()) → Excel 2026, GRID 2026 |
| MONTH | 2 | skip 2 | `functions/function-now.xlsx!A1` =MONTH(NOW()) → Excel 11, GRID 11 |
