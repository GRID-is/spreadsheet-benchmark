# hyperformula vs hyperformula@3.5.0 (hyperformula)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-lacks-function 241, volatile 10. Fails where the engine's error came from another cell: 2. Fails where both answered an error, but a different one: 78. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (22): ARRAYFORMULA, ARRAY_CONSTRAIN, COUNTUNIQUE, HF.ADD, HF.CONCAT, HF.DIVIDE, HF.EQ, HF.GT, HF.GTE, HF.LT, HF.LTE, HF.MINUS, HF.MULTIPLY, HF.NE, HF.POW, HF.UMINUS, HF.UNARY_PERCENT, HF.UPLUS, INTERVAL, ISBINARY, MAXPOOL, MEDIANPOOL. Never-supported functions present: ARRAY_CONSTRAIN.
Upstream literal differs from Excel's value in 677 of 4491 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 4166 | 3331 | 594 | 0 | 0 | 241 | 84.9% |
| other | 576 | 527 | 39 | 0 | 0 | 10 | 93.1% |
| **total** | 4742 | 3858 | 633 | 0 | 0 | 251 | 85.9% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 4742 | 3858 | 633 | 0 | 0 | 251 | 85.9% |
| **total** | 4742 | 3858 | 633 | 0 | 0 | 251 | 85.9% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| DATEDIF | 89 | fail 89 | `functions/function-datedif.xlsx!A3` =DATEDIF(1.9, 1.8, "Y") → Excel 0, HyperFormula "#NUM!" |
| YEARFRAC | 36 | fail 36 | `functions/function-yearfrac.xlsx!A5` =YEARFRAC("31/03/2020", "30/03/2020") → Excel "#VALUE!", HyperFormula 0 |
| TEXT | 34 | fail 34 | `functions/function-text.xlsx!B1` =TEXT(A1, "mm/dd/yyyy") → Excel "01/02/1900", HyperFormula "01/01/1900" |
| WORKDAY.INTL | 29 | fail 29 | `functions/function-workday.intl.xlsx!A3` =WORKDAY.INTL(0, 1, "1010102") → Excel "#VALUE!", HyperFormula "#NUM!" |
| WEEKNUM | 27 | fail 27 | `functions/function-weeknum.xlsx!A4` =WEEKNUM("02/08/2020", "21") → Excel 6, HyperFormula 31 |
| XLOOKUP | 20 | fail 20 | `functions/function-xlookup.xlsx!A3` =XLOOKUP(1, B1:B2, C1:C2, 0, 0.5) → Excel 0, HyperFormula "#VALUE!" |
| HF.EQ | 17 | skip 17 | `functions/function-hfeq.xlsx!A1` =HF.EQ(1) → Excel "#NAME?", HyperFormula "#N/A" |
| HF.GT | 17 | skip 17 | `functions/function-hfgt.xlsx!A1` =HF.GT(1) → Excel "#NAME?", HyperFormula "#N/A" |
| HF.GTE | 17 | skip 17 | `functions/function-hfgte.xlsx!A1` =HF.GTE(1) → Excel "#NAME?", HyperFormula "#N/A" |
| HF.LT | 17 | skip 17 | `functions/function-hflt.xlsx!A1` =HF.LT(1) → Excel "#NAME?", HyperFormula "#N/A" |
| HF.LTE | 17 | skip 17 | `functions/function-hflte.xlsx!A1` =HF.LTE(1) → Excel "#NAME?", HyperFormula "#N/A" |
| HF.NE | 17 | skip 17 | `functions/function-hfne.xlsx!A1` =HF.NE(1) → Excel "#NAME?", HyperFormula "#N/A" |
| DAYS360 | 16 | fail 16 | `functions/function-days360.xlsx!A5` =DAYS360("31/03/2020", "30/03/2020") → Excel "#VALUE!", HyperFormula 0 |
| WEEKDAY | 16 | fail 16 | `functions/function-weekday.xlsx!A3` =WEEKDAY("31/07/2020", "2") → Excel "#VALUE!", HyperFormula 5 |
| NETWORKDAYS.INTL | 15 | fail 15 | `functions/function-networkdays.intl.xlsx!A3` =NETWORKDAYS.INTL(0, 1, "1010102") → Excel "#VALUE!", HyperFormula "#NUM!" |
| SUMIF | 15 | fail 15 | `functions/function-sumif.xlsx!A1` =SUMIF(B1:B2, "><foo", C1:C2) → Excel 0, HyperFormula "#VALUE!" |
| INTERVAL | 13 | skip 13 | `functions/function-interval.xlsx!B1` =INTERVAL("12/30/2018") → Excel "#NAME?", HyperFormula "#VALUE!" |
| HF.DIVIDE | 12 | skip 12 | `functions/function-hfdivide.xlsx!A1` =HF.DIVIDE(1) → Excel "#NAME?", HyperFormula "#N/A" |
| COUNTUNIQUE | 11 | skip 11 | `functions/function-countunique.xlsx!A1` =COUNTUNIQUE() → Excel "#NAME?", HyperFormula "#N/A" |
| DAY | 11 | fail 9, skip 2 | `functions/function-day.xlsx!B1` =DAY("12/30/2018") → Excel 30, HyperFormula "#VALUE!" |
| HF.ADD | 11 | skip 11 | `functions/function-hfadd.xlsx!A1` =HF.ADD(1) → Excel "#NAME?", HyperFormula "#N/A" |
| HF.MINUS | 11 | skip 11 | `functions/function-hfminus.xlsx!A1` =HF.MINUS(1) → Excel "#NAME?", HyperFormula "#N/A" |
| HF.MULTIPLY | 11 | skip 11 | `functions/function-hfmultiply.xlsx!A1` =HF.MULTIPLY(1) → Excel "#NAME?", HyperFormula "#N/A" |
| HF.UMINUS | 11 | skip 11 | `functions/function-hfuminus.xlsx!A1` =HF.UMINUS() → Excel "#NAME?", HyperFormula "#N/A" |
| DAYS | 10 | fail 10 | `functions/function-days.xlsx!A3` =DAYS(2, "12/30/2018") → Excel -43462, HyperFormula "#VALUE!" |
