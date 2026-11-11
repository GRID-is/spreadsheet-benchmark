# gnumeric vs gnumeric@1.12.62 (gnumeric)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, environment 3, excel-read-as-user-function 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 0. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 10 of 945 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 950 | 917 | 27 | 1 | 0 | 5 | 97.0% |
| overview | 650 | 646 | 0 | 0 | 0 | 4 | 100.0% |
| yalta2008 | 369 | 361 | 8 | 0 | 0 | 0 | 97.8% |
| **total** | 1969 | 1924 | 35 | 1 | 0 | 9 | 98.2% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1969 | 1924 | 35 | 1 | 0 | 9 | 98.2% |
| **total** | 1969 | 1924 | 35 | 1 | 0 | 9 | 98.2% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| IMPOWER | 3 | fail 3 | `functions/engfuns.xls!F41` =IMPOWER("8-2j",9.2) → Excel "-169691027.479611-208555947.133373j", Gnumeric "-169691027.47961047-208555947.13337198i" |
| IMSQRT | 3 | fail 3 | `functions/engfuns.xls!F45` =IMSQRT("5-9j") → Excel "2.76546832751588-1.62721082545978j", Gnumeric "2.765468327515884-1.6272108254597806j" |
| IMSIN | 3 | fail 3 | `functions/engfuns.xls!F44` =IMSIN("5-9j") → Excel "-3885.12199806271-1149.26923044064j", Gnumeric "-3885.1219980627084-1149.2692304406407j" |
| IMEXP | 3 | fail 3 | `functions/engfuns.xls!F37` =IMEXP("9-2j") → Excel "-3372.07274273649-7368.11336469684j", Gnumeric "-3372.07274273649-7368.113364696836j" |
| IMLN | 3 | fail 3 | `functions/engfuns.xls!F38` =IMLN("3.3-9.2j") → Excel "2.27972016714096-1.2263959247521j", Gnumeric "2.2797201671409586-1.2263959247520995j" |
| IMLOG10 | 3 | fail 3 | `functions/engfuns.xls!F39` =IMLOG10("3.3-9.2j") → Excel "0.990069888872877-0.532616982748472j", Gnumeric "0.9900698888728773-0.5326169827484725j" |
| IMLOG2 | 3 | fail 3 | `functions/engfuns.xls!F40` =IMLOG2("3.3-9.2j") → Excel "3.28894097974882-1.76931531880629j", Gnumeric "3.2889409797488196-1.7693153188062882j" |
| IMCOS | 3 | fail 3 | `functions/engfuns.xls!F35` =IMCOS("5-9j") → Excel "1149.26926544734-3885.12187972205j", Gnumeric "1149.2692654473356-3885.1218797220517j" |
| INFO | 3 | skip 3 | `functions/infofuns.xls!F29` =INFO("osversion") → Excel "Windows (64-bit) NT 10.00", Gnumeric "Darwin version 24.6.0" |
| IF | 3 | fail 3 | `yalta2008/yalta2008.xls!J10` =IF(ISERROR(F10),FALSE,ABS(H10-F10)<=I10) → Excel false, Gnumeric true |
| AND | 2 | fail 2 | `yalta2008/yalta2008.xls!B1` =AND(J10:J65536) → Excel false, Gnumeric true |
| _XLUDF.MINIFS | 2 | skip 2 | `overview/ifs-funcs.xlsx!B22` =_xludf.MINIFS($C$6:$C$12,$B$6:$B$12,">=2") → Excel "#NAME?", Gnumeric "#NAME?" |
| _XLUDF.MAXIFS | 2 | skip 2 | `overview/ifs-funcs.xlsx!B30` =_xludf.MAXIFS($C$6:$C$12,$B$6:$B$12,"<>5") → Excel "#NAME?", Gnumeric "#NAME?" |
| CONVERT | 1 | fail 1 | `functions/engfuns.xls!F21` =CONVERT(7.9,"cal","J") → Excel 33.07572, Gnumeric 33.07567923492655 |
| IMDIV | 1 | fail 1 | `functions/engfuns.xls!D36` =IMDIV("3-j","4-j") → Excel "0.764705882352941-0.0588235294117647j", Gnumeric "0.7647058823529411-0.058823529411764705j" |
| BESSELJ | 1 | fail 1 | `functions/engfuns.xls!D14` =BESSELJ(8,3) → Excel -0.29112524185253685, Gnumeric -0.2911322070659522 |
| NETWORKDAYS | 1 | unsupported 1 | `functions/datefuns.xls!F30` =NETWORKDAYS("1999/01/22","2001/05/21",@{"2000/03/23";"2000/0 → Excel 606, Gnumeric "#NAME?" |
| NOW | 1 | skip 1 | `functions/datefuns.xls!B31` =NOW() → Excel 46330.76197175926, Gnumeric 46335.139403791014 |
| TODAY | 1 | skip 1 | `functions/datefuns.xls!B35` =TODAY() → Excel 46330, Gnumeric 46335 |
| CHIINV | 1 | fail 1 | `yalta2008/yalta2008.xls!D20` =CHIINV(A20,B20) → Excel 50.84417133244917, Gnumeric 50.844127911818155 |
