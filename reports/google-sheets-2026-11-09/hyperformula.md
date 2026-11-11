# hyperformula vs google-sheets@2026-11-09 (hyperformula)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-lacks-function 241, volatile 10. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 95. Fails where the case's own cell matches but its spill does not: 3.
Functions outside Excel's list (22): ARRAYFORMULA, ARRAY_CONSTRAIN, COUNTUNIQUE, HF.ADD, HF.CONCAT, HF.DIVIDE, HF.EQ, HF.GT, HF.GTE, HF.LT, HF.LTE, HF.MINUS, HF.MULTIPLY, HF.NE, HF.POW, HF.UMINUS, HF.UNARY_PERCENT, HF.UPLUS, INTERVAL, ISBINARY, MAXPOOL, MEDIANPOOL. Never-supported functions present: ARRAY_CONSTRAIN.
Upstream literal differs from Excel's value in 677 of 4491 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 4166 | 3529 | 331 | 63 | 2 | 241 | 89.9% |
| other | 576 | 545 | 21 | 0 | 0 | 10 | 96.3% |
| **total** | 4742 | 4074 | 352 | 63 | 2 | 251 | 90.7% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 4742 | 4074 | 352 | 63 | 2 | 251 | 90.7% |
| **total** | 4742 | 4074 | 352 | 63 | 2 | 251 | 90.7% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| XLOOKUP | 28 | fail 28 | `functions/function-xlookup.xlsx!A6` =XLOOKUP(1, B1:E1, B2:C3) → Excel "#VALUE!", Google Sheets "#N/A" |
| TEXT | 25 | fail 25 | `functions/function-text.xlsx!B1` =TEXT(A1, "mm/dd/yyyy") → Excel "01/02/1900", Google Sheets "01/01/1900" |
| HF.EQ | 17 | skip 17 | `functions/function-hfeq.xlsx!A1` =HF.EQ(1) → Excel "#NAME?", Google Sheets "#NAME?" |
| HF.GT | 17 | skip 17 | `functions/function-hfgt.xlsx!A1` =HF.GT(1) → Excel "#NAME?", Google Sheets "#NAME?" |
| HF.GTE | 17 | skip 17 | `functions/function-hfgte.xlsx!A1` =HF.GTE(1) → Excel "#NAME?", Google Sheets "#NAME?" |
| HF.LT | 17 | skip 17 | `functions/function-hflt.xlsx!A1` =HF.LT(1) → Excel "#NAME?", Google Sheets "#NAME?" |
| HF.LTE | 17 | skip 17 | `functions/function-hflte.xlsx!A1` =HF.LTE(1) → Excel "#NAME?", Google Sheets "#NAME?" |
| HF.NE | 17 | skip 17 | `functions/function-hfne.xlsx!A1` =HF.NE(1) → Excel "#NAME?", Google Sheets "#NAME?" |
| BESSELI | 15 | unsupported 15 | `functions/function-besseli.xlsx!A1` =BESSELI("foo", 1) → Excel "#VALUE!", Google Sheets "#NAME?" |
| BESSELJ | 15 | unsupported 15 | `functions/function-besselj.xlsx!A1` =BESSELJ("foo", 1) → Excel "#VALUE!", Google Sheets "#NAME?" |
| BESSELK | 15 | unsupported 15 | `functions/function-besselk.xlsx!A1` =BESSELK("foo", 1) → Excel "#VALUE!", Google Sheets "#NAME?" |
| BESSELY | 15 | unsupported 15 | `functions/function-bessely.xlsx!A1` =BESSELY("foo", 1) → Excel "#VALUE!", Google Sheets "#NAME?" |
| INDEX | 13 | fail 13 | `functions/function-index.xlsx!A2` =INDEX(B1:D3, 1, 1, 42) → Excel "#REF!", Google Sheets "" |
| INTERVAL | 13 | skip 13 | `functions/function-interval.xlsx!B1` =INTERVAL("12/30/2018") → Excel "#NAME?", Google Sheets "#NAME?" |
| WORKDAY.INTL | 13 | fail 13 | `functions/function-workday.intl.xlsx!A3` =WORKDAY.INTL(0, 1, "1010102") → Excel "#VALUE!", Google Sheets "#NUM!" |
| HF.DIVIDE | 12 | skip 12 | `functions/function-hfdivide.xlsx!A1` =HF.DIVIDE(1) → Excel "#NAME?", Google Sheets "#NAME?" |
| COUNTUNIQUE | 11 | skip 11 | `functions/function-countunique.xlsx!A1` =COUNTUNIQUE() → Excel "#NAME?", Google Sheets "#N/A" |
| HF.ADD | 11 | skip 11 | `functions/function-hfadd.xlsx!A1` =HF.ADD(1) → Excel "#NAME?", Google Sheets "#NAME?" |
| HF.MINUS | 11 | skip 11 | `functions/function-hfminus.xlsx!A1` =HF.MINUS(1) → Excel "#NAME?", Google Sheets "#NAME?" |
| HF.MULTIPLY | 11 | skip 11 | `functions/function-hfmultiply.xlsx!A1` =HF.MULTIPLY(1) → Excel "#NAME?", Google Sheets "#NAME?" |
| HF.UMINUS | 11 | skip 11 | `functions/function-hfuminus.xlsx!A1` =HF.UMINUS() → Excel "#NAME?", Google Sheets "#NAME?" |
| NETWORKDAYS.INTL | 11 | fail 11 | `functions/function-networkdays.intl.xlsx!A3` =NETWORKDAYS.INTL(0, 1, "1010102") → Excel "#VALUE!", Google Sheets "#NUM!" |
| HF.POW | 10 | skip 10 | `functions/function-hfpow.xlsx!A1` =HF.POW(1) → Excel "#NAME?", Google Sheets "#NAME?" |
| HF.UPLUS | 10 | skip 10 | `functions/function-hfuplus.xlsx!A1` =HF.UPLUS() → Excel "#NAME?", Google Sheets "#NAME?" |
| HF.CONCAT | 9 | skip 9 | `functions/function-hfconcat.xlsx!A1` =HF.CONCAT(1) → Excel "#NAME?", Google Sheets "#NAME?" |

## Errors

- `functions/function-mirr.xlsx`: Google Sheets failed to export sheet 4 should return #DIV 0! if any
- `functions/function-mirr.xlsx`: Google Sheets failed to export sheet 4 should return #DIV 0! if any
