# fast-formula-parser vs fast-formula-parser@1.0.19 (fast-formula-parser)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 5. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 4. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 70 of 1465 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| date | 158 | 152 | 3 | 0 | 0 | 3 | 98.1% |
| engineering | 376 | 345 | 31 | 0 | 0 | 0 | 91.8% |
| information | 58 | 56 | 2 | 0 | 0 | 0 | 96.6% |
| logical | 37 | 37 | 0 | 0 | 0 | 0 | 100.0% |
| math | 272 | 266 | 4 | 0 | 0 | 2 | 98.5% |
| operators | 56 | 56 | 0 | 0 | 0 | 0 | 100.0% |
| reference | 88 | 79 | 9 | 0 | 0 | 0 | 89.8% |
| statistical | 262 | 254 | 8 | 0 | 0 | 0 | 96.9% |
| text | 81 | 74 | 6 | 1 | 0 | 0 | 91.4% |
| trigonometry | 82 | 82 | 0 | 0 | 0 | 0 | 100.0% |
| **total** | 1470 | 1401 | 63 | 1 | 0 | 5 | 95.6% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1470 | 1401 | 63 | 1 | 0 | 5 | 95.6% |
| **total** | 1470 | 1401 | 63 | 1 | 0 | 5 | 95.6% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| HYPGEOM.DIST | 5 | fail 5 | `statistical.xlsx!AA129` =HYPGEOM.DIST(1,4,0,20,TRUE) → Excel 1, fast-formula-parser "#NUM!" |
| IMCOSH | 3 | fail 3 | `engineering.xlsx!AA237` =IMCOSH("4+3j") → Excel "-27.0349456030742+3.85115333481178j", fast-formula-parser "-27.034945603074224+3.851153334811777j" |
| IMCOS | 3 | fail 3 | `engineering.xlsx!AA229` =IMCOS(1) → Excel "0.54030230586814", fast-formula-parser "0.5403023058681398" |
| ROWS | 3 | fail 3 | `reference.xlsx!AA68` =ROWS({1,2,3}) → Excel 1, fast-formula-parser "#VALUE!" |
| COLUMNS | 3 | fail 3 | `reference.xlsx!AA22` =COLUMNS({1,2,3}) → Excel 3, fast-formula-parser "#VALUE!" |
| ASC | 3 | fail 3 | `text.xlsx!AA2` =ASC("ヲァィゥ") → Excel "ヲァィゥ", fast-formula-parser "ｦｧｨｩ" |
| DBCS | 3 | fail 3 | `text.xlsx!AA22` =DBCS("ｦｧｨｩ") → Excel "ｦｧｨｩ", fast-formula-parser "ヲァィゥ" |
| YEAR | 2 | skip 2 | `date.xlsx!AA103` =YEAR(NOW()) → Excel 2026, fast-formula-parser 2026 |
| IMLOG10 | 2 | fail 2 | `engineering.xlsx!AA262` =IMLOG10("3+4i") → Excel "0.698970004336019+0.402719196273373i", fast-formula-parser "0.6989700043360187+0.4027191962733731i" |
| IMLOG2 | 2 | fail 2 | `engineering.xlsx!AA267` =IMLOG2("3+4j") → Excel "2.32192809488736+1.33780421245098j", fast-formula-parser "2.321928094887362+1.3378042124509761j" |
| IMSINH | 2 | fail 2 | `engineering.xlsx!AA304` =IMSINH("4+3j") → Excel "-27.0168132580039+3.85373803791938j", fast-formula-parser "-27.016813258003932+3.853738037919377j" |
| IMSQRT | 2 | fail 2 | `engineering.xlsx!AA307` =IMSQRT("1+j") → Excel "1.09868411346781+0.455089860562227j", fast-formula-parser "1.0986841134678098+0.45508986056222733j" |
| IMSIN | 2 | fail 2 | `engineering.xlsx!AA299` =IMSIN("4+3i") → Excel "-7.61923172032141-6.548120040911i", fast-formula-parser "-7.61923172032141-6.5481200409110025i" |
| IMLN | 2 | fail 2 | `engineering.xlsx!AA260` =IMLN("3+4j") → Excel "1.6094379124341+0.927295218001612j", fast-formula-parser "1.6094379124341003+0.9272952180016122j" |
| IMEXP | 2 | fail 2 | `engineering.xlsx!AA256` =IMEXP("1+i") → Excel "1.46869393991589+2.28735528717884i", fast-formula-parser "1.4686939399158851+2.2873552871788423i" |
| FREQUENCY | 2 | fail 2 | `statistical.xlsx!AA85` =FREQUENCY({1,2,3,4,5},{5,1,11}) → Excel 4, fast-formula-parser 1 |
| YEARFRAC | 1 | fail 1 | `date.xlsx!AA158` =YEARFRAC("2012/1/1","2012/7/30", 5) → Excel "#NUM!", fast-formula-parser "#VALUE!" |
| HOUR | 1 | skip 1 | `date.xlsx!AA104` =HOUR(NOW()) → Excel 18, fast-formula-parser 2 |
| DAY | 1 | fail 1 | `date.xlsx!AA60` =DAY(-12) → Excel "#NUM!", fast-formula-parser "#VALUE!" |
| DATEVALUE | 1 | fail 1 | `date.xlsx!AA39` =DATEVALUE("Wednesday, December 3, 2014") → Excel "#VALUE!", fast-formula-parser 41976 |
| OCT2HEX | 1 | fail 1 | `engineering.xlsx!AA371` =OCT2HEX(100, 0) → Excel "#NUM!", fast-formula-parser "40" |
| OCT2DEC | 1 | fail 1 | `engineering.xlsx!AA347` =OCT2DEC(TRUE) → Excel "#VALUE!", fast-formula-parser "#NUM!" |
| IMTAN | 1 | fail 1 | `engineering.xlsx!AA316` =IMTAN("4+3i") → Excel "0.00490825806749606+1.00070953606723i", fast-formula-parser "0.004908258067495992+1.000709536067233i" |
| IMPOWER | 1 | fail 1 | `engineering.xlsx!AA268` =IMPOWER("2+3i", 3) → Excel "-46+9.00000000000001i", fast-formula-parser "-45.99999999999999+9.000000000000007i" |
| IMSUB | 1 | fail 1 | `engineering.xlsx!AA310` =IMSUB("13+4j","5+3i") → Excel "#NUM!", fast-formula-parser "#VALUE!" |
