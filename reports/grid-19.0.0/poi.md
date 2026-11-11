# poi vs Apiary 19.0.0 (e018d83eb8c6a829)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 0. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 47 of 3178 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| eval | 2453 | 2451 | 0 | 0 | 0 | 2 | 100.0% |
| functions | 693 | 691 | 0 | 0 | 0 | 2 | 100.0% |
| matrix | 34 | 34 | 0 | 0 | 0 | 0 | 100.0% |
| multisheet | 2 | 2 | 0 | 0 | 0 | 0 | 100.0% |
| **total** | 3182 | 3178 | 0 | 0 | 0 | 4 | 100.0% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 3182 | 3178 | 0 | 0 | 0 | 4 | 100.0% |
| **total** | 3182 | 3178 | 0 | 0 | 0 | 4 | 100.0% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| WEEKNUM | 4 | skip 4 | `eval/FormulaEvalTestData.xls!J1524` =WEEKNUM("99/9/2008",-1) → Excel "#VALUE!", GRID "#VALUE!" |
