# fast-formula-parser vs Apiary 19.0.0 (e018d83eb8c6a829)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 5. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 0. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 70 of 1465 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| date | 158 | 155 | 0 | 0 | 0 | 3 | 100.0% |
| engineering | 376 | 376 | 0 | 0 | 0 | 0 | 100.0% |
| information | 58 | 58 | 0 | 0 | 0 | 0 | 100.0% |
| logical | 37 | 37 | 0 | 0 | 0 | 0 | 100.0% |
| math | 272 | 270 | 0 | 0 | 0 | 2 | 100.0% |
| operators | 56 | 56 | 0 | 0 | 0 | 0 | 100.0% |
| reference | 88 | 88 | 0 | 0 | 0 | 0 | 100.0% |
| statistical | 262 | 262 | 0 | 0 | 0 | 0 | 100.0% |
| text | 81 | 81 | 0 | 0 | 0 | 0 | 100.0% |
| trigonometry | 82 | 82 | 0 | 0 | 0 | 0 | 100.0% |
| **total** | 1470 | 1465 | 0 | 0 | 0 | 5 | 100.0% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1470 | 1465 | 0 | 0 | 0 | 5 | 100.0% |
| **total** | 1470 | 1465 | 0 | 0 | 0 | 5 | 100.0% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| YEAR | 2 | skip 2 | `date.xlsx!AA103` =YEAR(NOW()) → Excel 2026, GRID 2026 |
| HOUR | 1 | skip 1 | `date.xlsx!AA104` =HOUR(NOW()) → Excel 18, GRID 1 |
| RAND | 1 | skip 1 | `math.xlsx!AA200` =RAND() > 0 → Excel true, GRID true |
| RANDBETWEEN | 1 | skip 1 | `math.xlsx!AA201` =RANDBETWEEN(-1,1) >= -1 → Excel true, GRID true |
