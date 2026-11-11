# fast-formula-parser vs mog@1.0.0 (fast-formula-parser)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 5. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 49. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 70 of 1465 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| date | 158 | 137 | 18 | 0 | 0 | 3 | 88.4% |
| engineering | 376 | 316 | 60 | 0 | 0 | 0 | 84.0% |
| information | 58 | 57 | 0 | 1 | 0 | 0 | 98.3% |
| logical | 37 | 31 | 6 | 0 | 0 | 0 | 83.8% |
| math | 272 | 245 | 25 | 0 | 0 | 2 | 90.7% |
| operators | 56 | 56 | 0 | 0 | 0 | 0 | 100.0% |
| reference | 88 | 75 | 12 | 1 | 0 | 0 | 85.2% |
| statistical | 262 | 223 | 39 | 0 | 0 | 0 | 85.1% |
| text | 81 | 71 | 10 | 0 | 0 | 0 | 87.7% |
| trigonometry | 82 | 74 | 8 | 0 | 0 | 0 | 90.2% |
| **total** | 1470 | 1285 | 178 | 2 | 0 | 5 | 87.7% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1470 | 1285 | 178 | 2 | 0 | 5 | 87.7% |
| **total** | 1470 | 1285 | 178 | 2 | 0 | 5 | 87.7% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| SUM | 9 | fail 8, unsupported 1 | `math.xlsx!AA239` =SUM("1", {"1"},) → Excel 1, Mog 0 |
| DATEVALUE | 6 | fail 6 | `date.xlsx!AA45` =DATEVALUE("12/3") → Excel 46359, Mog "#VALUE!" |
| AVERAGEA | 4 | fail 4 | `statistical.xlsx!AA227` =AVERAGEA({TRUE,1,2,"12"}, 3) → Excel 1.5, Mog 3.8 |
| NUMBERVALUE | 4 | fail 4 | `text.xlsx!AA56` =NUMBERVALUE("3.5%",,) → Excel "#VALUE!", Mog 0.035 |
| WORKDAY.INTL | 3 | fail 3 | `date.xlsx!AA130` =WORKDAY.INTL(DATE(2012,1,1),30,"1") → Excel "#VALUE!", Mog 40949 |
| IMSUB | 3 | fail 3 | `engineering.xlsx!AA311` =IMSUB(TRUE,"5+3i") → Excel "#VALUE!", Mog "#NUM!" |
| IMTAN | 3 | fail 3 | `engineering.xlsx!AA316` =IMTAN("4+3i") → Excel "0.00490825806749606+1.00070953606723i", Mog "0.00490825806749599+1.00070953606723i" |
| IMDIV | 3 | fail 3 | `engineering.xlsx!AA250` =IMDIV(TRUE,"52") → Excel "#VALUE!", Mog "#NUM!" |
| IMCOT | 3 | fail 3 | `engineering.xlsx!AA239` =IMCOT("4+3i") → Excel "0.00490118239430447-0.999266927805902i", Mog "0.0049011823943044-0.999266927805902i" |
| BITRSHIFT | 3 | fail 3 | `engineering.xlsx!AA108` =BITRSHIFT(9999999999,-50) → Excel 0, Mog "#NUM!" |
| DELTA | 3 | fail 3 | `engineering.xlsx!AA160` =DELTA(5,FALSE) → Excel "#VALUE!", Mog 0 |
| LCM | 3 | fail 3 | `math.xlsx!AA137` =LCM(5, 2, (A1, A2)) → Excel "#VALUE!", Mog 100000 |
| VLOOKUP | 3 | fail 3 | `reference.xlsx!AA81` =VLOOKUP("b", {"c";TRUE;"abc";65;63;61;"b";"a";TRUE;FALSE}, 1 → Excel "b", Mog "a" |
| HLOOKUP | 3 | fail 3 | `reference.xlsx!AA38` =HLOOKUP(3, 3, 1,FALSE) → Excel "#N/A", Mog "#VALUE!" |
| AVERAGE | 3 | fail 3 | `statistical.xlsx!AA226` =AVERAGE(TRUE, "3a") → Excel "#VALUE!", Mog 1 |
| AVEDEV | 3 | fail 3 | `statistical.xlsx!AA220` =AVEDEV({4,1,6,7,5,4,3},(A7,B7),"1") → Excel 1.8, Mog 1.7037037037037037 |
| T.INV.2T | 3 | fail 3 | `statistical.xlsx!AA206` =T.INV.2T(0.546449, 60.1) → Excel 0.606533075825759, Mog 0.6065273191294683 |
| F.INV.RT | 3 | fail 3 | `statistical.xlsx!AA64` =F.INV.RT(0.01, 1000000000000, 4) → Excel "#NUM!", Mog 13.457758547821571 |
| ASC | 3 | fail 3 | `text.xlsx!AA2` =ASC("ヲァィゥ") → Excel "ヲァィゥ", Mog "ｦｧｨｩ" |
| DBCS | 3 | fail 3 | `text.xlsx!AA22` =DBCS("ｦｧｨｩ") → Excel "ｦｧｨｩ", Mog "ヲァィゥ" |
| NETWORKDAYS.INTL | 2 | fail 2 | `date.xlsx!AA101` =NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31), "1") → Excel "#VALUE!", Mog 22 |
| YEARFRAC | 2 | fail 2 | `date.xlsx!AA147` =YEARFRAC("2020/2/5","2021/2/1", 1) → Excel 0.9890710382513661, Mog 0.9904240766073872 |
| HOUR | 2 | fail 1, skip 1 | `date.xlsx!AA104` =HOUR(NOW()) → Excel 18, Mog 5 |
| YEAR | 2 | skip 2 | `date.xlsx!AA103` =YEAR(NOW()) → Excel 2026, Mog 2026 |
| ISOWEEKNUM | 2 | fail 2 | `date.xlsx!AA121` =ISOWEEKNUM("3/9/2012") → Excel 10, Mog 11 |
