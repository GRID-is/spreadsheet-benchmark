# hyperformula vs spread-sheets@19.2.5 (hyperformula)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-lacks-function 241, volatile 10. Fails where the engine's error came from another cell: 5. Fails where both answered an error, but a different one: 66. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (22): ARRAYFORMULA, ARRAY_CONSTRAIN, COUNTUNIQUE, HF.ADD, HF.CONCAT, HF.DIVIDE, HF.EQ, HF.GT, HF.GTE, HF.LT, HF.LTE, HF.MINUS, HF.MULTIPLY, HF.NE, HF.POW, HF.UMINUS, HF.UNARY_PERCENT, HF.UPLUS, INTERVAL, ISBINARY, MAXPOOL, MEDIANPOOL. Never-supported functions present: ARRAY_CONSTRAIN.
Upstream literal differs from Excel's value in 677 of 4491 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 4166 | 3568 | 355 | 2 | 0 | 241 | 90.9% |
| other | 576 | 532 | 34 | 0 | 0 | 10 | 94.0% |
| **total** | 4742 | 4100 | 389 | 2 | 0 | 251 | 91.3% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 4742 | 4100 | 389 | 2 | 0 | 251 | 91.3% |
| **total** | 4742 | 4100 | 389 | 2 | 0 | 251 | 91.3% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| TEXT | 26 | fail 26 | `functions/function-text.xlsx!B1` =TEXT(A1, "mm/dd/yyyy") → Excel "01/02/1900", SpreadJS "01/01/1900" |
| HF.EQ | 17 | skip 17 | `functions/function-hfeq.xlsx!A1` =HF.EQ(1) → Excel "#NAME?", SpreadJS "#NAME?" |
| HF.GT | 17 | skip 17 | `functions/function-hfgt.xlsx!A1` =HF.GT(1) → Excel "#NAME?", SpreadJS "#NAME?" |
| HF.GTE | 17 | skip 17 | `functions/function-hfgte.xlsx!A1` =HF.GTE(1) → Excel "#NAME?", SpreadJS "#NAME?" |
| HF.LT | 17 | skip 17 | `functions/function-hflt.xlsx!A1` =HF.LT(1) → Excel "#NAME?", SpreadJS "#NAME?" |
| HF.LTE | 17 | skip 17 | `functions/function-hflte.xlsx!A1` =HF.LTE(1) → Excel "#NAME?", SpreadJS "#NAME?" |
| HF.NE | 17 | skip 17 | `functions/function-hfne.xlsx!A1` =HF.NE(1) → Excel "#NAME?", SpreadJS "#NAME?" |
| XLOOKUP | 15 | fail 15 | `functions/function-xlookup.xlsx!A5` =XLOOKUP(1, B1:B2, C1:C2, 0, B1:B2) → Excel "#VALUE!", SpreadJS 0 |
| INTERVAL | 13 | skip 13 | `functions/function-interval.xlsx!B1` =INTERVAL("12/30/2018") → Excel "#NAME?", SpreadJS "#NAME?" |
| HF.DIVIDE | 12 | skip 12 | `functions/function-hfdivide.xlsx!A1` =HF.DIVIDE(1) → Excel "#NAME?", SpreadJS "#NAME?" |
| COUNTUNIQUE | 11 | skip 11 | `functions/function-countunique.xlsx!A1` =COUNTUNIQUE() → Excel "#NAME?", SpreadJS "#NAME?" |
| HF.ADD | 11 | skip 11 | `functions/function-hfadd.xlsx!A1` =HF.ADD(1) → Excel "#NAME?", SpreadJS "#NAME?" |
| HF.MINUS | 11 | skip 11 | `functions/function-hfminus.xlsx!A1` =HF.MINUS(1) → Excel "#NAME?", SpreadJS "#NAME?" |
| HF.MULTIPLY | 11 | skip 11 | `functions/function-hfmultiply.xlsx!A1` =HF.MULTIPLY(1) → Excel "#NAME?", SpreadJS "#NAME?" |
| HF.UMINUS | 11 | skip 11 | `functions/function-hfuminus.xlsx!A1` =HF.UMINUS() → Excel "#NAME?", SpreadJS "#NAME?" |
| HF.POW | 10 | skip 10 | `functions/function-hfpow.xlsx!A1` =HF.POW(1) → Excel "#NAME?", SpreadJS "#NAME?" |
| HF.UPLUS | 10 | skip 10 | `functions/function-hfuplus.xlsx!A1` =HF.UPLUS() → Excel "#NAME?", SpreadJS "#NAME?" |
| SUMPRODUCT | 10 | fail 10 | `functions/function-sumproduct.xlsx!A2` =SUMPRODUCT(A1,B1) → Excel 0, SpreadJS 42 |
| HF.CONCAT | 9 | skip 9 | `functions/function-hfconcat.xlsx!A1` =HF.CONCAT(1) → Excel "#NAME?", SpreadJS "#NAME?" |
| HF.UNARY_PERCENT | 9 | skip 9 | `functions/function-hfunary_percent.xlsx!A1` =HF.UNARY_PERCENT() → Excel "#NAME?", SpreadJS "#NAME?" |
| DAY | 8 | fail 6, skip 2 | `functions/function-day.xlsx!A1` =DAY(0) → Excel 0, SpreadJS 30 |
| DATE | 7 | fail 7 | `functions/function-date.xlsx!A1` =DATE(1900, 1, 1) → Excel 1, SpreadJS 2 |
| VLOOKUP | 7 | fail 7 | `functions/function-vlookup.xlsx!A1` =VLOOKUP(1, A2:B3, 3) → Excel "#N/A", SpreadJS "#REF!" |
| XNPV | 7 | fail 7 | `functions/function-xnpv.xlsx!A1` =XNPV(-1, A2:D2, A3:D3) → Excel "#NUM!", SpreadJS null |
| ISBINARY | 6 | skip 6 | `functions/function-isbinary.xlsx!A1` =ISBINARY("1010") → Excel "#NAME?", SpreadJS "#NAME?" |
