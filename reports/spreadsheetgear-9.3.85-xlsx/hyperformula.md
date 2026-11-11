# hyperformula vs spreadsheetgear@9.3.85 (hyperformula)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-lacks-function 241, volatile 10. Fails where the engine's error came from another cell: 3. Fails where both answered an error, but a different one: 21. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (22): ARRAYFORMULA, ARRAY_CONSTRAIN, COUNTUNIQUE, HF.ADD, HF.CONCAT, HF.DIVIDE, HF.EQ, HF.GT, HF.GTE, HF.LT, HF.LTE, HF.MINUS, HF.MULTIPLY, HF.NE, HF.POW, HF.UMINUS, HF.UNARY_PERCENT, HF.UPLUS, INTERVAL, ISBINARY, MAXPOOL, MEDIANPOOL. Never-supported functions present: ARRAY_CONSTRAIN.
Upstream literal differs from Excel's value in 677 of 4491 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 4166 | 3613 | 90 | 222 | 0 | 241 | 92.1% |
| other | 576 | 564 | 2 | 0 | 0 | 10 | 99.6% |
| **total** | 4742 | 4177 | 92 | 222 | 0 | 251 | 93.0% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 4742 | 4177 | 92 | 222 | 0 | 251 | 93.0% |
| **total** | 4742 | 4177 | 92 | 222 | 0 | 251 | 93.0% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| XLOOKUP | 139 | unsupported 139 | `functions/function-xlookup.xlsx!A6` =XLOOKUP(1, B1:E1, B2:C3) → Excel "#VALUE!", SpreadsheetGear "#NAME?" |
| MAXIFS | 24 | unsupported 24 | `functions/function-maxifs.xlsx!A1` =MAXIFS(B1:B2, C1:C2, "><foo") → Excel 0, SpreadsheetGear "#NAME?" |
| MINIFS | 24 | unsupported 24 | `functions/function-minifs.xlsx!A1` =MINIFS(B1:B2, C1:C2, "><foo") → Excel 0, SpreadsheetGear "#NAME?" |
| HF.EQ | 17 | skip 17 | `functions/function-hfeq.xlsx!A1` =HF.EQ(1) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| HF.GT | 17 | skip 17 | `functions/function-hfgt.xlsx!A1` =HF.GT(1) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| HF.GTE | 17 | skip 17 | `functions/function-hfgte.xlsx!A1` =HF.GTE(1) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| HF.LT | 17 | skip 17 | `functions/function-hflt.xlsx!A1` =HF.LT(1) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| HF.LTE | 17 | skip 17 | `functions/function-hflte.xlsx!A1` =HF.LTE(1) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| HF.NE | 17 | skip 17 | `functions/function-hfne.xlsx!A1` =HF.NE(1) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| INTERVAL | 13 | skip 13 | `functions/function-interval.xlsx!B1` =INTERVAL("12/30/2018") → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| HF.DIVIDE | 12 | skip 12 | `functions/function-hfdivide.xlsx!A1` =HF.DIVIDE(1) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| COUNTUNIQUE | 11 | skip 11 | `functions/function-countunique.xlsx!A1` =COUNTUNIQUE() → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| HF.ADD | 11 | skip 11 | `functions/function-hfadd.xlsx!A1` =HF.ADD(1) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| HF.MINUS | 11 | skip 11 | `functions/function-hfminus.xlsx!A1` =HF.MINUS(1) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| HF.MULTIPLY | 11 | skip 11 | `functions/function-hfmultiply.xlsx!A1` =HF.MULTIPLY(1) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| HF.UMINUS | 11 | skip 11 | `functions/function-hfuminus.xlsx!A1` =HF.UMINUS() → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| HF.POW | 10 | skip 10 | `functions/function-hfpow.xlsx!A1` =HF.POW(1) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| HF.UPLUS | 10 | skip 10 | `functions/function-hfuplus.xlsx!A1` =HF.UPLUS() → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| SWITCH | 10 | unsupported 10 | `functions/function-switch.xlsx!A1` =SWITCH(1,2,3) → Excel "#N/A", SpreadsheetGear "#NAME?" |
| HF.CONCAT | 9 | skip 9 | `functions/function-hfconcat.xlsx!A1` =HF.CONCAT(1) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| HF.UNARY_PERCENT | 9 | skip 9 | `functions/function-hfunary_percent.xlsx!A1` =HF.UNARY_PERCENT() → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| IFS | 7 | unsupported 7 | `functions/function-ifs.xlsx!B2` =IFS(A2>30, "A", A2>20, "B", A2>10, "C") → Excel "B", SpreadsheetGear "#NAME?" |
| ISBINARY | 6 | skip 6 | `functions/function-isbinary.xlsx!A1` =ISBINARY("1010") → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| MAXPOOL | 5 | skip 5 | `other/matrix-plugin.xlsx!A4` =maxpool(A1:F3,3) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| MEDIANPOOL | 5 | skip 5 | `other/matrix-plugin.xlsx!A3` =medianpool(A1:F2,2) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
