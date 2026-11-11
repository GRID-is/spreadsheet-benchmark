# hyperformula vs ironcalc@0.9.1 (hyperformula)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-lacks-function 45, volatile 10. Fails where the engine's error came from another cell: 4. Fails where both answered an error, but a different one: 102. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (7): ARRAYFORMULA, ARRAY_CONSTRAIN, COUNTUNIQUE, INTERVAL, ISBINARY, MAXPOOL, MEDIANPOOL. Never-supported functions present: ARRAY_CONSTRAIN.
Upstream literal differs from Excel's value in 873 of 4687 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 4166 | 3312 | 536 | 0 | 273 | 45 | 80.4% |
| other | 576 | 468 | 48 | 0 | 50 | 10 | 82.7% |
| **total** | 4742 | 3780 | 584 | 0 | 323 | 55 | 80.6% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 4742 | 3780 | 584 | 0 | 323 | 55 | 80.6% |
| **total** | 4742 | 3780 | 584 | 0 | 323 | 55 | 80.6% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| DATEDIF | 90 | fail 90 | `functions/function-datedif.xlsx!A1` =DATEDIF(1, 2, "abcd") → Excel "#NUM!", IronCalc "#VALUE!" |
| TEXT | 61 | fail 61 | `other/date-and-time-arithmetic.xlsx!C32` =TEXT(A32-B32, "hh:mm") → Excel "01:31", IronCalc "01:30" |
| SUMIF | 23 | fail 23 | `functions/function-sumif.xlsx!A2` =SUMIF(B1, ">0", B2:D2) → Excel 0, IronCalc "#VALUE!" |
| HF.EQ | 17 | error 17 | `functions/function-hfeq.xlsx!A1` =HF.EQ(1) → Excel "#NAME?", IronCalc null |
| HF.GT | 17 | error 17 | `functions/function-hfgt.xlsx!A1` =HF.GT(1) → Excel "#NAME?", IronCalc null |
| HF.GTE | 17 | error 17 | `functions/function-hfgte.xlsx!A1` =HF.GTE(1) → Excel "#NAME?", IronCalc null |
| HF.LT | 17 | error 17 | `functions/function-hflt.xlsx!A1` =HF.LT(1) → Excel "#NAME?", IronCalc null |
| HF.LTE | 17 | error 17 | `functions/function-hflte.xlsx!A1` =HF.LTE(1) → Excel "#NAME?", IronCalc null |
| HF.NE | 17 | error 17 | `functions/function-hfne.xlsx!A1` =HF.NE(1) → Excel "#NAME?", IronCalc null |
| VLOOKUP | 17 | fail 17 | `functions/function-vlookup.xlsx!A1` =VLOOKUP(D1:E1, A2:B3, 2, TRUE()) → Excel "#VALUE!", IronCalc "#N/A" |
| REPLACE | 14 | error 14 | `functions/function-replace.xlsx!A1` =REPLACE("foobar", 2, 2, "uu") → Excel "fuubar", IronCalc null |
| INTERVAL | 13 | skip 13 | `functions/function-interval.xlsx!B1` =INTERVAL("12/30/2018") → Excel "#NAME?", IronCalc "#NAME?" |
| NETWORKDAYS.INTL | 13 | fail 13 | `functions/function-networkdays.intl.xlsx!A5` =NETWORKDAYS.INTL(0, 1, "1111111") → Excel 0, IronCalc "#NUM!" |
| XNPV | 13 | error 13 | `functions/function-xnpv.xlsx!A1` =XNPV(-1, A2:D2, A3:D3) → Excel "#NUM!", IronCalc null |
| ADDRESS | 12 | fail 12 | `functions/function-address.xlsx!A7` =ADDRESS(-1, -1, 4, FALSE()) → Excel "R[-1]C[-1]", IronCalc "#VALUE!" |
| HF.DIVIDE | 12 | error 12 | `functions/function-hfdivide.xlsx!A1` =HF.DIVIDE(1) → Excel "#NAME?", IronCalc null |
| XLOOKUP | 12 | fail 12 | `functions/function-xlookup.xlsx!A9` =XLOOKUP(1, B1:C2, D3:E4) → Excel "#VALUE!", IronCalc "#ERROR!" |
| COUNTUNIQUE | 11 | skip 11 | `functions/function-countunique.xlsx!A1` =COUNTUNIQUE() → Excel "#NAME?", IronCalc "#NAME?" |
| HF.ADD | 11 | error 11 | `functions/function-hfadd.xlsx!A1` =HF.ADD(1) → Excel "#NAME?", IronCalc null |
| HF.MINUS | 11 | error 11 | `functions/function-hfminus.xlsx!A1` =HF.MINUS(1) → Excel "#NAME?", IronCalc null |
| HF.MULTIPLY | 11 | error 11 | `functions/function-hfmultiply.xlsx!A1` =HF.MULTIPLY(1) → Excel "#NAME?", IronCalc null |
| HF.UMINUS | 11 | error 11 | `functions/function-hfuminus.xlsx!A1` =HF.UMINUS() → Excel "#NAME?", IronCalc null |
| MID | 11 | error 11 | `functions/function-mid.xlsx!A1` =MID("foo", 1, 2) → Excel "fo", IronCalc null |
| MIRR | 11 | error 11 | `functions/function-mirr.xlsx!A4` =MIRR(B4:D4, 0.2, 0.1) → Excel -0.2614510541240036, IronCalc null |
| NETWORKDAYS | 11 | fail 11 | `functions/function-networkdays.xlsx!A3` =NETWORKDAYS(0, 6.9) → Excel 5, IronCalc "#NUM!" |

## Errors

- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `other/error-literals.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-column.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-column.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-column.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-column.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-column.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-column.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-column.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-column.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfadd.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfadd.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfadd.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfadd.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfadd.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfadd.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfadd.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfadd.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfadd.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfadd.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfadd.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfconcat.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfconcat.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfconcat.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfconcat.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfconcat.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfconcat.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfconcat.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfconcat.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfconcat.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfdivide.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfdivide.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfdivide.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfdivide.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfdivide.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfdivide.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfdivide.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfdivide.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfdivide.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfdivide.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfdivide.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfdivide.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfeq.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfeq.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfeq.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfeq.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfeq.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfeq.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfeq.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfeq.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfeq.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfeq.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfeq.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfeq.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfeq.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfeq.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfeq.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfeq.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfeq.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfgte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflt.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hflte.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfmultiply.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfmultiply.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfmultiply.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfmultiply.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfmultiply.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfmultiply.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfmultiply.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfmultiply.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfmultiply.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfmultiply.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfmultiply.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfne.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfne.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfne.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfne.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfne.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfne.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfne.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfne.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfne.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfne.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfne.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfne.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfne.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfne.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfne.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfne.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfne.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfpow.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfpow.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfpow.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfpow.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfpow.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfpow.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfpow.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfpow.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfpow.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfpow.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuminus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfunary_percent.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfunary_percent.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfunary_percent.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfunary_percent.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfunary_percent.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfunary_percent.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfunary_percent.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfunary_percent.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfunary_percent.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuplus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuplus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuplus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuplus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuplus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuplus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuplus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuplus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuplus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-hfuplus.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isformula.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isformula.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isformula.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isformula.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isformula.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isformula.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isna.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isna.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isna.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isna.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isna.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isna.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isref.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isref.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isref.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isref.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isref.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isref.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isref.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-isref.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mid.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mid.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mid.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mid.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mid.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mid.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mid.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mid.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mid.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mid.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mid.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mirr.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mirr.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mirr.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mirr.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mirr.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mirr.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mirr.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mirr.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mirr.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mirr.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-mirr.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-replace.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-replace.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-replace.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-replace.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-replace.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-replace.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-replace.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-replace.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-replace.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-replace.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-replace.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-replace.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-replace.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-replace.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-xnpv.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-xnpv.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-xnpv.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-xnpv.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-xnpv.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-xnpv.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-xnpv.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-xnpv.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-xnpv.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-xnpv.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-xnpv.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-xnpv.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `functions/function-xnpv.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
