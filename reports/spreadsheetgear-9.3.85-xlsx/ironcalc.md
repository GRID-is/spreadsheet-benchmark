# ironcalc vs spreadsheetgear@9.3.85 (ironcalc)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, excel-lacks-function 4. Fails where the engine's error came from another cell: 3. Fails where both answered an error, but a different one: 116. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (3): GELU, HELLO, XYZ. Never-supported functions present: none.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| (root) | 708 | 667 | 2 | 39 | 0 | 0 | 94.2% |
| DATABASE | 566 | 541 | 25 | 0 | 0 | 0 | 95.6% |
| DATE_AND_TIME | 2623 | 2559 | 41 | 10 | 10 | 3 | 97.7% |
| ENGINEERING | 2976 | 2950 | 26 | 0 | 0 | 0 | 99.1% |
| FINANCIAL | 3086 | 2879 | 207 | 0 | 0 | 0 | 93.3% |
| INFORMATION | 409 | 251 | 6 | 23 | 129 | 0 | 61.4% |
| LOGICAL | 533 | 465 | 0 | 59 | 8 | 1 | 87.4% |
| LOOKUP_AND_REFERENCE | 1047 | 407 | 5 | 531 | 104 | 0 | 38.9% |
| MATH_AND_TRIGONOMETRY | 56100 | 55389 | 506 | 127 | 76 | 2 | 98.7% |
| STATISTICAL | 3184 | 2796 | 388 | 0 | 0 | 0 | 87.8% |
| TEXT | 1457 | 1148 | 9 | 257 | 43 | 0 | 78.8% |
| **total** | 72689 | 70052 | 1215 | 1046 | 370 | 6 | 96.4% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 72689 | 70052 | 1215 | 1046 | 370 | 6 | 96.4% |
| **total** | 72689 | 70052 | 1215 | 1046 | 370 | 6 | 96.4% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| XLOOKUP | 439 | unsupported 439 | `LOOKUP_AND_REFERENCE/XLOOKUP.xlsx!H6` =XLOOKUP(E2,A:A,B:B,,12) → Excel "#VALUE!", SpreadsheetGear "#NAME?" |
| COUNTIF | 252 | fail 252 | `STATISTICAL/COUNTIF.xlsx!AJ12` =COUNTIF(AG$2:AG$15,CONCAT("<>", AH12)) → Excel 8, SpreadsheetGear 0 |
| ACCRINT | 150 | fail 150 | `FINANCIAL/ACCRINT.xlsx!B36` =ACCRINT(DATE(2016,6,1), DATE(2017,12,1), DATE(2016,9,1), 0.3 → Excel 239.25000000000003, SpreadsheetGear 252.99999999999997 |
| COUNTIFS | 104 | fail 104 | `STATISTICAL/COUNTIFS.xlsx!AE12` =COUNTIFS(AB$2:AB$15,CONCAT("<>", AC12)) → Excel 8, SpreadsheetGear 0 |
| CEILING.MATH | 79 | fail 78, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AG136` =CEILING.MATH($A136,AG$2) → Excel "#VALUE!", SpreadsheetGear 0 |
| TEXTAFTER | 71 | unsupported 71 | `TEXT/TEXTBEFORE_TEXTAFTER.xlsx!G40` =TEXTAFTER(A40,B40,C40,D40,E40,F40) → Excel "#VALUE!", SpreadsheetGear "#NAME?" |
| TEXTBEFORE | 70 | unsupported 70 | `TEXT/TEXTBEFORE_TEXTAFTER.xlsx!J39` =TEXTBEFORE(A39,B39) → Excel "#N/A", SpreadsheetGear "#NAME?" |
| CEILING | 62 | fail 61, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AG43` =CEILING($A43,AG$2) → Excel "#VALUE!", SpreadsheetGear 0 |
| FLOOR.MATH | 61 | fail 60, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI127` =FLOOR.MATH($A127) → Excel "#VALUE!", SpreadsheetGear 45301 |
| FLOOR.PRECISE | 61 | fail 60, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI80` =FLOOR.PRECISE($A80) → Excel "#VALUE!", SpreadsheetGear 45301 |
| ISO.CEILING | 61 | fail 60, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI205` =ISO.CEILING($A205) → Excel "#VALUE!", SpreadsheetGear 45301 |
| CEILING.PRECISE | 61 | fail 60, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AI80` =CEILING.PRECISE($A80) → Excel "#VALUE!", SpreadsheetGear 45301 |
| FLOOR | 60 | fail 59, unsupported 1 | `MATH_AND_TRIGONOMETRY/FLOOR_CEILING.xlsx!AG33` =FLOOR($A33,$AD$47:$AH$47) → Excel "#VALUE!", SpreadsheetGear 45300 |
| MAXIFS | 50 | unsupported 50 | `MATH_AND_TRIGONOMETRY/SUMIFS.xlsx!H32` =MAXIFS($A$25:$A$40,$B$25:$B$40,C32,$B$25:$B$40,D32) → Excel 2, SpreadsheetGear "#NAME?" |
| MINIFS | 50 | unsupported 50 | `MATH_AND_TRIGONOMETRY/SUMIFS.xlsx!H18` =MINIFS($A$2:$A$20,$B$2:$B$20,C18) → Excel 0, SpreadsheetGear "#NAME?" |
| IFS | 49 | unsupported 49 | `LOGICAL/IFS.xlsx!H12` =IFS(A12,B12,F12,G12) → Excel "#VALUE!", SpreadsheetGear "#NAME?" |
| MAX | 47 | error 47 | `INFORMATION/ISREF.xlsx!D5` =MAX(FILTER(L1:L2, M1:M2, 0)) → Excel 0, SpreadsheetGear null |
| TEXTJOIN | 40 | unsupported 40 | `TEXT/TEXTJOIN.xlsx!I13` =TEXTJOIN(A13,B13,C13,D13:E13,F13) → Excel "Monday0Tuesday0Wednesday0Thursday", SpreadsheetGear "#NAME?" |
| INDIRECT | 35 | error 35 | `LOOKUP_AND_REFERENCE/INDIRECT.xlsx!E3` =INDIRECT("B2", 1) → Excel 5, SpreadsheetGear null |
| WEEKDAY | 33 | fail 30, error 3 | `DATE_AND_TIME/WEEKDAY_WEEKNUM_ISOWEEKNUM.xlsx!M18` =WEEKDAY($A18,M$1) → Excel "#NUM!", SpreadsheetGear "#VALUE!" |
| SUM | 30 | fail 1, unsupported 10, error 19 | `INFORMATION/ISREF.xlsx!D14` =SUM(FILTER(L1:L2, M1:M2, L1:L2)) → Excel 30, SpreadsheetGear null |
| CONCAT | 29 | unsupported 29 | `TEXT/CONCAT.xlsx!E7` =CONCAT(A7,B7,C7) → Excel "TRUE", SpreadsheetGear "#NAME?" |
| SWITCH | 23 | unsupported 23 | `logical.xlsx!I6` =SWITCH(H6, "Monday", 1, "Tuesday",2,"WEDNESDAY",3) → Excel "#N/A", SpreadsheetGear "#NAME?" |
| ISREF | 22 | error 22 | `INFORMATION/ISREF.xlsx!D20` =ISREF(FILTER(L1:L2, L1:L2>0)) → Excel false, SpreadsheetGear null |
| TEXTSPLIT | 19 | unsupported 19 | `TEXT/TEXTSPLIT.xlsx!B20` =TEXTSPLIT(A20,",",";",TRUE,0,"-") → Excel "1", SpreadsheetGear "#NAME?" |

## Errors

- `DATE_AND_TIME/EOMONTH_issue_1388.xlsx`: spreadsheetgear: System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `DATE_AND_TIME/EOMONTH_issue_1388.xlsx`: spreadsheetgear: System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `DATE_AND_TIME/EOMONTH_issue_1388.xlsx`: spreadsheetgear: System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `DATE_AND_TIME/EOMONTH_issue_1388.xlsx`: spreadsheetgear: System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `DATE_AND_TIME/EOMONTH_issue_1388.xlsx`: spreadsheetgear: System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `DATE_AND_TIME/EOMONTH_issue_1388.xlsx`: spreadsheetgear: System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `DATE_AND_TIME/EOMONTH_issue_1388.xlsx`: spreadsheetgear: System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `DATE_AND_TIME/EOMONTH_issue_1388.xlsx`: spreadsheetgear: System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `DATE_AND_TIME/EOMONTH_issue_1388.xlsx`: spreadsheetgear: System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `DATE_AND_TIME/EOMONTH_issue_1388.xlsx`: spreadsheetgear: System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `INFORMATION/ISOMITTED.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISOMITTED.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISOMITTED.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/ISREF.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `INFORMATION/error_broadcast.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `LOGICAL/LAMBDA.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `LOGICAL/LAMBDA.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `LOGICAL/LAMBDA.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `LOGICAL/LAMBDA.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `LOGICAL/LAMBDA.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `LOGICAL/LAMBDA.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `LOGICAL/LAMBDA.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `LOGICAL/LAMBDA.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `LOOKUP_AND_REFERENCE/INDIRECT.xlsx`: timeout after 55000 ms
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MAX_MIN.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `MATH_AND_TRIGONOMETRY/MIN_MAX_extras.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `TEXT/MINUS_MINUS.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
