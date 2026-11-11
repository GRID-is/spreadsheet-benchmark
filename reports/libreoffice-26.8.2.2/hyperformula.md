# hyperformula vs libreoffice@26.8.2.2 (hyperformula)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-lacks-function 241, volatile 10. Fails where the engine's error came from another cell: 4. Fails where both answered an error, but a different one: 366. Fails where the case's own cell matches but its spill does not: 1.
Functions outside Excel's list (22): ARRAYFORMULA, ARRAY_CONSTRAIN, COUNTUNIQUE, HF.ADD, HF.CONCAT, HF.DIVIDE, HF.EQ, HF.GT, HF.GTE, HF.LT, HF.LTE, HF.MINUS, HF.MULTIPLY, HF.NE, HF.POW, HF.UMINUS, HF.UNARY_PERCENT, HF.UPLUS, INTERVAL, ISBINARY, MAXPOOL, MEDIANPOOL. Never-supported functions present: ARRAY_CONSTRAIN.
Upstream literal differs from Excel's value in 677 of 4491 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 4166 | 3218 | 682 | 25 | 0 | 241 | 82.0% |
| other | 576 | 501 | 65 | 0 | 0 | 10 | 88.5% |
| **total** | 4742 | 3719 | 747 | 25 | 0 | 251 | 82.8% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 4742 | 3719 | 747 | 25 | 0 | 251 | 82.8% |
| **total** | 4742 | 3719 | 747 | 25 | 0 | 251 | 82.8% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| TEXT | 27 | fail 27 | `functions/function-text.xlsx!B1` =TEXT(A1, "mm/dd/yyyy") → Excel "01/02/1900", LibreOffice "01/01/1900" |
| HF.EQ | 17 | skip 17 | `functions/function-hfeq.xlsx!A1` =HF.EQ(1) → Excel "#NAME?", LibreOffice "#NAME?" |
| HF.GT | 17 | skip 17 | `functions/function-hfgt.xlsx!A1` =HF.GT(1) → Excel "#NAME?", LibreOffice "#NAME?" |
| HF.GTE | 17 | skip 17 | `functions/function-hfgte.xlsx!A1` =HF.GTE(1) → Excel "#NAME?", LibreOffice "#NAME?" |
| HF.LT | 17 | skip 17 | `functions/function-hflt.xlsx!A1` =HF.LT(1) → Excel "#NAME?", LibreOffice "#NAME?" |
| HF.LTE | 17 | skip 17 | `functions/function-hflte.xlsx!A1` =HF.LTE(1) → Excel "#NAME?", LibreOffice "#NAME?" |
| HF.NE | 17 | skip 17 | `functions/function-hfne.xlsx!A1` =HF.NE(1) → Excel "#NAME?", LibreOffice "#NAME?" |
| XLOOKUP | 14 | fail 14 | `functions/function-xlookup.xlsx!A4` =XLOOKUP(1, B1:B2, C1:C2, 0, "string") → Excel "#VALUE!", LibreOffice 0 |
| INTERVAL | 13 | skip 13 | `functions/function-interval.xlsx!B1` =INTERVAL("12/30/2018") → Excel "#NAME?", LibreOffice "#NAME?" |
| HF.DIVIDE | 12 | skip 12 | `functions/function-hfdivide.xlsx!A1` =HF.DIVIDE(1) → Excel "#NAME?", LibreOffice "#NAME?" |
| COUNTUNIQUE | 11 | skip 11 | `functions/function-countunique.xlsx!A1` =COUNTUNIQUE() → Excel "#NAME?", LibreOffice "#NAME?" |
| HF.ADD | 11 | skip 11 | `functions/function-hfadd.xlsx!A1` =HF.ADD(1) → Excel "#NAME?", LibreOffice "#NAME?" |
| HF.MINUS | 11 | skip 11 | `functions/function-hfminus.xlsx!A1` =HF.MINUS(1) → Excel "#NAME?", LibreOffice "#NAME?" |
| HF.MULTIPLY | 11 | skip 11 | `functions/function-hfmultiply.xlsx!A1` =HF.MULTIPLY(1) → Excel "#NAME?", LibreOffice "#NAME?" |
| HF.UMINUS | 11 | skip 11 | `functions/function-hfuminus.xlsx!A1` =HF.UMINUS() → Excel "#NAME?", LibreOffice "#NAME?" |
| HEX2OCT | 10 | fail 10 | `functions/function-hex2oct.xlsx!A1` =HEX2OCT("foo") → Excel "#NUM!", LibreOffice "#VALUE!" |
| HF.POW | 10 | skip 10 | `functions/function-hfpow.xlsx!A1` =HF.POW(1) → Excel "#NAME?", LibreOffice "#NAME?" |
| HF.UPLUS | 10 | skip 10 | `functions/function-hfuplus.xlsx!A1` =HF.UPLUS() → Excel "#NAME?", LibreOffice "#NAME?" |
| OCT2BIN | 10 | fail 10 | `functions/function-oct2bin.xlsx!A1` =OCT2BIN("foo") → Excel "#NUM!", LibreOffice "#VALUE!" |
| SMALL | 10 | fail 10 | `functions/function-small.xlsx!A1` =SMALL(A2:D2,0) → Excel "#NUM!", LibreOffice "#VALUE!" |
| VLOOKUP | 10 | fail 10 | `functions/function-vlookup.xlsx!A1` =VLOOKUP(1, A2:B3, 3) → Excel "#N/A", LibreOffice "#VALUE!" |
| TRUE | 9 | fail 9 | `other/boolean-operators.xlsx!B1` ="foo"<TRUE() → Excel true, LibreOffice false |
| ARABIC | 9 | fail 9 | `functions/function-arabic.xlsx!A1` =ARABIC(" XD ") → Excel 490, LibreOffice "#VALUE!" |
| HEX2BIN | 9 | fail 9 | `functions/function-hex2bin.xlsx!A1` =HEX2BIN("foo") → Excel "#NUM!", LibreOffice "#VALUE!" |
| HF.CONCAT | 9 | skip 9 | `functions/function-hfconcat.xlsx!A1` =HF.CONCAT(1) → Excel "#NAME?", LibreOffice "#NAME?" |
