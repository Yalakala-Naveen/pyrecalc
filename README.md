# pyrecalc — Recalc

**Evaluate Excel formulas in Python, without Excel — and verify the
workbooks your code generates.**

openpyxl writes formulas but never evaluates them. So every workbook your
pipeline generates has values nobody has checked until a human opens it
in Excel — and if the cached values in the file are stale, the workbook
is *silently wrong*. The existing pure-Python evaluators are fragmented:
`formulas` is EUPL-1.1 copyleft, `pycel` is GPL-3.0, `xlcalculator` has
been inactive for over a year. pyrecalc is the MIT-licensed,
verification-first evaluator.

```python
from pyrecalc import Workbook

wb = Workbook("model.xlsx")
wb.set("Inputs!B3", 0.12)      # plain A1 addressing, no bracket-key strings
print(wb.value("Outputs!C2"))  # -> the computed value, in-process

report = wb.verify()           # every formula cell vs its cached value
print(report.ok, len(report.mismatches))
```

```bash
pyrecalc check model.xlsx      # exit 0 = all cached values match, 1 = mismatch/error
pyrecalc value model.xlsx "Outputs!C2"
```

## What v1 evaluates

- Arithmetic, comparison and concatenation operators, percent, unary minus
- Cell refs (relative and `$`-absolute), sheet-qualified refs
  (`Inputs!B3`, `'My Sheet'!B3`), ranges (`A1:B5`)
- SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, IF, IFERROR, ROUND, ROUNDUP,
  ROUNDDOWN, INT, MOD, ABS, SQRT, POWER, AND, OR, NOT, CONCAT,
  CONCATENATE, LEN, UPPER, LOWER, LEFT, RIGHT, MID, TRIM, PROPER
- Criteria aggregation: SUMIF, COUNTIF, AVERAGEIF (Excel criteria
  semantics — operator prefixes like `">=20"` and `"<>apple"`, text
  compared case-insensitively), SUMIFS, COUNTIFS, AVERAGEIFS and
  SUMPRODUCT
- Lookups: VLOOKUP (exact and approximate), INDEX and MATCH —
  including the INDEX+MATCH combination generated workbooks use
  instead of VLOOKUP
- Excel error values as values — `#DIV/0!`, `#VALUE!`, `#REF!`,
  `#NAME?`, `#N/A` propagate through dependents exactly like Excel;
  circular references are detected and reported (`#CIRC!`), never a hang

## Honest scope

Deterministic evaluation only — no model in the path. Out of scope for
v1, stated plainly: array formulas, defined names, external-workbook
references, DATABASE/CUBE/WEB functions, and XLOOKUP (VLOOKUP,
INDEX and MATCH are in — see above). Dates pass through as the
numbers Excel stores.

## Why

The incumbent evaluators together are downloaded ≈5.2M times a month
(pypistats, Nov 2026) despite the copyleft licenses and stalled
maintenance — and a $49 packaged workaround exists just to check
generated workbooks. The wedge pyrecalc serves first is that check:
*does this workbook's own arithmetic agree with the values saved in it?*

## Development

```bash
python -m unittest discover -s tests   # 17 tests, generated-workbook fixtures
```

MIT © Naveen Yalakala
