# fast-formula-parser vs spreadsheetgear@9.3.85 (fast-formula-parser)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 5. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 5. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 70 of 1465 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| date | 158 | 152 | 3 | 0 | 0 | 3 | 98.1% |
| engineering | 376 | 364 | 12 | 0 | 0 | 0 | 96.8% |
| information | 58 | 58 | 0 | 0 | 0 | 0 | 100.0% |
| logical | 37 | 33 | 0 | 4 | 0 | 0 | 89.2% |
| math | 272 | 261 | 9 | 0 | 0 | 2 | 96.7% |
| operators | 56 | 56 | 0 | 0 | 0 | 0 | 100.0% |
| reference | 88 | 86 | 2 | 0 | 0 | 0 | 97.7% |
| statistical | 262 | 247 | 11 | 4 | 0 | 0 | 94.3% |
| text | 81 | 67 | 1 | 13 | 0 | 0 | 82.7% |
| trigonometry | 82 | 81 | 1 | 0 | 0 | 0 | 98.8% |
| **total** | 1470 | 1405 | 39 | 21 | 0 | 5 | 95.9% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1470 | 1405 | 39 | 21 | 0 | 5 | 95.9% |
| **total** | 1470 | 1405 | 39 | 21 | 0 | 5 | 95.9% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| LCM | 5 | fail 5 | `math.xlsx!AA137` =LCM(5, 2, (A1, A2)) → Excel "#VALUE!", SpreadsheetGear 100000 |
| HYPGEOM.DIST | 5 | fail 5 | `statistical.xlsx!AA129` =HYPGEOM.DIST(1,4,0,20,TRUE) → Excel 1, SpreadsheetGear "#NUM!" |
| CONCAT | 5 | unsupported 5 | `text.xlsx!AA12` =CONCAT(0, {1,2,3;5,6,7}) → Excel "0123567", SpreadsheetGear "#NAME?" |
| IFS | 4 | unsupported 4 | `logical.xlsx!AA15` =IFS(D5<60,"F",D5<70,"D",D5<80,"C",D5<90,"B",D5>=90,"A") → Excel "F", SpreadsheetGear "#NAME?" |
| FORECAST.LINEAR | 4 | unsupported 4 | `statistical.xlsx!AA83` =FORECAST.LINEAR(30,{6,6,4},{2,1}) → Excel "#N/A", SpreadsheetGear "#NAME?" |
| BESSELJ | 3 | fail 3 | `engineering.xlsx!AA13` =BESSELJ(1.9, {2,3}) → Excel 0.3299258286697852, SpreadsheetGear 0.3299257276923872 |
| ASC | 3 | unsupported 3 | `text.xlsx!AA2` =ASC("ヲァィゥ") → Excel "ヲァィゥ", SpreadsheetGear "#NAME?" |
| DBCS | 3 | unsupported 3 | `text.xlsx!AA22` =DBCS("ｦｧｨｩ") → Excel "ｦｧｨｩ", SpreadsheetGear "#NAME?" |
| NETWORKDAYS.INTL | 2 | fail 2 | `date.xlsx!AA100` =NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"0010001",{"2 → Excel 20, SpreadsheetGear 22 |
| YEAR | 2 | skip 2 | `date.xlsx!AA103` =YEAR(NOW()) → Excel 2026, SpreadsheetGear 2026 |
| BITRSHIFT | 2 | fail 2 | `engineering.xlsx!AA100` =BITRSHIFT(13,2.7) → Excel 3, SpreadsheetGear "#NUM!" |
| GCD | 2 | fail 2 | `math.xlsx!AA117` =GCD(5, 2, {3,"7a"}) → Excel "#VALUE!", SpreadsheetGear 1 |
| STANDARDIZE | 2 | fail 2 | `statistical.xlsx!AA186` =STANDARDIZE(42, 40, 0) → Excel "#NUM!", SpreadsheetGear "#DIV/0!" |
| BAHTTEXT | 2 | unsupported 2 | `text.xlsx!AA4` =BAHTTEXT(63147.89) → Excel "หกหมื่นสามพันหนึ่งร้อยสี่สิบเจ็ดบาทแปดสิบเก้าสตางค์", SpreadsheetGear "#NAME?" |
| HOUR | 1 | skip 1 | `date.xlsx!AA104` =HOUR(NOW()) → Excel 18, SpreadsheetGear 16 |
| WORKDAY | 1 | fail 1 | `date.xlsx!AA128` =WORKDAY(DATE(2008,10,1),151, {"2008/11/26","2008/12/4","2008 → Excel 39937, SpreadsheetGear 39933 |
| IMSUB | 1 | fail 1 | `engineering.xlsx!AA310` =IMSUB("13+4j","5+3i") → Excel "#NUM!", SpreadsheetGear "#VALUE!" |
| IMSEC | 1 | fail 1 | `engineering.xlsx!AA295` =IMSEC("4+3i") → Excel "-0.0652940278579471-0.0752249603027732i", SpreadsheetGear "-0.065294027857947-0.0752249603027732i" |
| IMDIV | 1 | fail 1 | `engineering.xlsx!AA253` =IMDIV("-238+240i","10+24j") → Excel "#NUM!", SpreadsheetGear "#VALUE!" |
| IMCOT | 1 | fail 1 | `engineering.xlsx!AA239` =IMCOT("4+3i") → Excel "0.00490118239430447-0.999266927805902i", SpreadsheetGear "0.00490118239430447-0.999266927805901i" |
| BITLSHIFT | 1 | fail 1 | `engineering.xlsx!AA84` =BITLSHIFT(4,2.4) → Excel 16, SpreadsheetGear "#NUM!" |
| BIN2OCT | 1 | fail 1 | `engineering.xlsx!AA58` =BIN2OCT({1100100}) → Excel "144", SpreadsheetGear 100 |
| BIN2HEX | 1 | fail 1 | `engineering.xlsx!AA48` =BIN2HEX(1111111111,{1}) → Excel "FFFFFFFFFF", SpreadsheetGear 100 |
| RAND | 1 | skip 1 | `math.xlsx!AA200` =RAND() > 0 → Excel true, SpreadsheetGear true |
| RANDBETWEEN | 1 | skip 1 | `math.xlsx!AA201` =RANDBETWEEN(-1,1) >= -1 → Excel true, SpreadsheetGear true |
