# gnumeric vs Apiary 19.0.0 (e018d83eb8c6a829)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, environment 3, excel-read-as-user-function 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 0. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 10 of 945 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 950 | 945 | 0 | 0 | 0 | 5 | 100.0% |
| overview | 650 | 646 | 0 | 0 | 0 | 4 | 100.0% |
| yalta2008 | 369 | 365 | 4 | 0 | 0 | 0 | 98.9% |
| **total** | 1969 | 1956 | 4 | 0 | 0 | 9 | 99.8% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1969 | 1956 | 4 | 0 | 0 | 9 | 99.8% |
| **total** | 1969 | 1956 | 4 | 0 | 0 | 9 | 99.8% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| INFO | 3 | skip 3 | `functions/infofuns.xls!F29` =INFO("osversion") → Excel "Windows (64-bit) NT 10.00", GRID "" |
| _XLUDF.MINIFS | 2 | skip 2 | `overview/ifs-funcs.xlsx!B22` =_xludf.MINIFS($C$6:$C$12,$B$6:$B$12,">=2") → Excel "#NAME?", GRID "#NAME?" |
| _XLUDF.MAXIFS | 2 | skip 2 | `overview/ifs-funcs.xlsx!B30` =_xludf.MAXIFS($C$6:$C$12,$B$6:$B$12,"<>5") → Excel "#NAME?", GRID "#NAME?" |
| NOW | 1 | skip 1 | `functions/datefuns.xls!B31` =NOW() → Excel 46330.76197175926, GRID 46335.06439702546 |
| TODAY | 1 | skip 1 | `functions/datefuns.xls!B35` =TODAY() → Excel 46330, GRID 46335 |
| IF | 1 | fail 1 | `yalta2008/yalta2008.xls!H20` =IF(ISERROR(D20),FALSE,ABS(F20-D20)<=G20) → Excel false, GRID true |
| CHIINV | 1 | fail 1 | `yalta2008/yalta2008.xls!D20` =CHIINV(A20,B20) → Excel 50.84417133244917, GRID 50.844127911818155 |
| AND | 1 | fail 1 | `yalta2008/yalta2008.xls!B1` =AND(H10:H65536) → Excel false, GRID true |
