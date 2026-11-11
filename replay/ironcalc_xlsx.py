#!/usr/bin/env python3
# Evaluate one fixture with IronCalc's own xlsx importer (the `ironcalc` PyPI
# package) and write the values of the requested cells as JSON. Driven by
# replay/ironcalc-xlsx.ts; not meant to be run by hand.
#
#   ironcalc_xlsx.py <fixture.xlsx> <targets.json> <values.json>
#
# targets.json: [{"sheet": "Sheet1", "ref": "A1"}, ...]
# values.json:  {"Sheet1!A1": <value>, ...} where a value is a number, a string,
#               a boolean, an error string such as "#DIV/0!", or null for an
#               empty cell or a sheet the workbook does not have. The values go
#               to a file because the importer prints its own diagnostics
#               ("Unexpected feature ...") to stdout.
import json
import re
import sys

import ironcalc

fixture, targets_path, values_path = sys.argv[1], sys.argv[2], sys.argv[3]
with open(targets_path) as f:
    targets = json.load(f)

model = ironcalc.load_from_xlsx(fixture, 'en', 'UTC')
model.evaluate()
sheet_index = {p['name']: i for i, p in enumerate(model.get_worksheets_properties())}


def coords(ref):
    m = re.match(r'^([A-Z]+)(\d+)$', ref)
    col = 0
    for ch in m.group(1):
        col = col * 26 + ord(ch) - 64
    return int(m.group(2)), col


out = {}
for t in targets:
    key = f"{t['sheet']}!{t['ref']}"
    sheet = sheet_index.get(t['sheet'])
    if sheet is None:
        out[key] = None
        continue
    row, col = coords(t['ref'])
    value = model.get_cell_value(sheet, row, col)
    if value == '' or value is None:
        out[key] = None
    else:
        out[key] = value

with open(values_path, 'w') as f:
    json.dump(out, f)
