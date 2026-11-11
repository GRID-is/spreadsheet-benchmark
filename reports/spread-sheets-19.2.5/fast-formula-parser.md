# fast-formula-parser vs spread-sheets@19.2.5 (fast-formula-parser)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 5. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 24. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 70 of 1465 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| date | 158 | 151 | 4 | 0 | 0 | 3 | 97.4% |
| engineering | 376 | 293 | 83 | 0 | 0 | 0 | 77.9% |
| information | 58 | 58 | 0 | 0 | 0 | 0 | 100.0% |
| logical | 37 | 31 | 6 | 0 | 0 | 0 | 83.8% |
| math | 272 | 253 | 17 | 0 | 0 | 2 | 93.7% |
| operators | 56 | 52 | 4 | 0 | 0 | 0 | 92.9% |
| reference | 88 | 82 | 6 | 0 | 0 | 0 | 93.2% |
| statistical | 262 | 229 | 29 | 4 | 0 | 0 | 87.4% |
| text | 81 | 62 | 7 | 12 | 0 | 0 | 76.5% |
| trigonometry | 82 | 74 | 8 | 0 | 0 | 0 | 90.2% |
| **total** | 1470 | 1285 | 164 | 16 | 0 | 5 | 87.7% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1470 | 1285 | 164 | 16 | 0 | 5 | 87.7% |
| **total** | 1470 | 1285 | 164 | 16 | 0 | 5 | 87.7% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| NUMBERVALUE | 12 | unsupported 12 | `text.xlsx!AA56` =NUMBERVALUE("3.5%",,) → Excel "#VALUE!", SpreadJS "#NAME?" |
| BITRSHIFT | 6 | fail 6 | `engineering.xlsx!AA105` =BITRSHIFT(9999999999,31) → Excel 4, SpreadJS 0 |
| LCM | 6 | fail 6 | `math.xlsx!AA137` =LCM(5, 2, (A1, A2)) → Excel "#VALUE!", SpreadJS 100000 |
| HYPGEOM.DIST | 5 | fail 5 | `statistical.xlsx!AA129` =HYPGEOM.DIST(1,4,0,20,TRUE) → Excel 1, SpreadJS "#NUM!" |
| IMSUB | 4 | fail 4 | `engineering.xlsx!AA311` =IMSUB(TRUE,"5+3i") → Excel "#VALUE!", SpreadJS "#NUM!" |
| IMCOS | 4 | fail 4 | `engineering.xlsx!AA229` =IMCOS(1) → Excel "0.54030230586814", SpreadJS "0.5403023058681398" |
| BITLSHIFT | 4 | fail 4 | `engineering.xlsx!AA82` =BITLSHIFT(9,-3) → Excel 1, SpreadJS 536870912 |
| BESSELJ | 4 | fail 4 | `engineering.xlsx!AA13` =BESSELJ(1.9, {2,3}) → Excel 0.3299258286697852, SpreadJS 0.32992572769238715 |
| FREQUENCY | 4 | fail 4 | `statistical.xlsx!AA84` =FREQUENCY({1,2,3,4,5},{1,5,11}) → Excel 1, SpreadJS "#SPILL!" |
| FORECAST.LINEAR | 4 | unsupported 4 | `statistical.xlsx!AA83` =FORECAST.LINEAR(30,{6,6,4},{2,1}) → Excel "#N/A", SpreadJS "#NAME?" |
| IMPOWER | 3 | fail 3 | `engineering.xlsx!AA270` =IMPOWER("2+3j", TRUE) → Excel "#VALUE!", SpreadJS "2+3i" |
| IMPRODUCT | 3 | fail 3 | `engineering.xlsx!AA272` =IMPRODUCT("3+4i","5-3j", "5-3j") → Excel "#VALUE!", SpreadJS "168-26i" |
| IMLOG2 | 3 | fail 3 | `engineering.xlsx!AA266` =IMLOG2(TRUE) → Excel "#VALUE!", SpreadJS "#NUM!" |
| IMLOG10 | 3 | fail 3 | `engineering.xlsx!AA262` =IMLOG10("3+4i") → Excel "0.698970004336019+0.402719196273373i", SpreadJS "0.6989700043360187+0.4027191962733731i" |
| IMSQRT | 3 | fail 3 | `engineering.xlsx!AA306` =IMSQRT(FALSE) → Excel "#VALUE!", SpreadJS "0" |
| IMSIN | 3 | fail 3 | `engineering.xlsx!AA299` =IMSIN("4+3i") → Excel "-7.61923172032141-6.548120040911i", SpreadJS "-7.61923172032141-6.5481200409110025i" |
| IMLN | 3 | fail 3 | `engineering.xlsx!AA260` =IMLN("3+4j") → Excel "1.6094379124341+0.927295218001612j", SpreadJS "1.6094379124341003+0.9272952180016122i" |
| IMEXP | 3 | fail 3 | `engineering.xlsx!AA256` =IMEXP("1+i") → Excel "1.46869393991589+2.28735528717884i", SpreadJS "1.4686939399158851+2.2873552871788423i" |
| IMDIV | 3 | fail 3 | `engineering.xlsx!AA250` =IMDIV(TRUE,"52") → Excel "#VALUE!", SpreadJS "#NUM!" |
| IMCOSH | 3 | fail 3 | `engineering.xlsx!AA237` =IMCOSH("4+3j") → Excel "-27.0349456030742+3.85115333481178j", SpreadJS "-27.034945603074224+3.851153334811777j" |
| DELTA | 3 | fail 3 | `engineering.xlsx!AA160` =DELTA(5,FALSE) → Excel "#VALUE!", SpreadJS 0 |
| GROWTH | 3 | fail 3 | `statistical.xlsx!AA118` =GROWTH({33100,47300,69000,102000,150000,220000}, {11,12,13,1 → Excel 724.7673986629499, SpreadJS "#SPILL!" |
| ASC | 3 | fail 3 | `text.xlsx!AA2` =ASC("ヲァィゥ") → Excel "ヲァィゥ", SpreadJS "ｦｧｨｩ" |
| DBCS | 3 | fail 3 | `text.xlsx!AA22` =DBCS("ｦｧｨｩ") → Excel "ｦｧｨｩ", SpreadJS "ヲァィゥ" |
| YEARFRAC | 2 | fail 2 | `date.xlsx!AA146` =YEARFRAC("1900/3/31","1900/1/30", 1) → Excel 0.16712328767123288, SpreadJS 0.163934426229508 |
