"""Produce tools/parity_expected.json: Python-side results for the parity battery."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
from pyrecalc.evaluator import Evaluator, iter_range  # noqa: E402
from pyrecalc.errors import ExcelError  # noqa: E402
from pyrecalc.parser import parse  # noqa: E402

GRID = {"A1": 10.0, "A2": 20.0, "A3": 30.0, "B1": 4.0, "B2": 5.0, "B3": 6.0,
        "C1": "apple", "C2": "banana", "C3": "apple", "D1": 2.5, "D2": -3.0, "D3": 0.0}

FORMULAS = [
    "A1+B1*2", "SUM(A1:A3)", "AVERAGE(B1:B3)", "MAX(A1:B3)-MIN(A1:B3)",
    "IF(A1>B1,\"big\",\"small\")", "SUMIF(C1:C3,\"apple\",A1:A3)",
    "COUNTIF(C1:C3,\"apple\")", "AVERAGEIF(C1:C3,\"apple\",A1:A3)",
    "SUMPRODUCT(A1:A3,B1:B3)", "ROUND(D1,0)", "ROUNDUP(D1,0)", "ROUNDDOWN(2.9,0)",
    "INT(-1.5)", "MOD(10,3)", "ABS(D2)", "SQRT(16)", "POWER(2,10)",
    "LEFT(C1,3)", "RIGHT(C2,4)", "MID(C1,2,3)", "LEN(C2)", "TRIM(\"  a  b  \")",
    "PROPER(\"hello world\")", "1/0", "IFERROR(1/0,\"caught\")", "COUNT(A1:B3)",
    "COUNTA(C1:C3)", "A1&\" items\"", "\"Apple\"=C1", "AND(A1>B1,B2>0)", "NOT(FALSE)",
    "50%", "SUM(A1:A3)/COUNT(A1:A3)", "NOSUCHFN(1)",
    "VLOOKUP(20,A1:C3,3,FALSE)", "VLOOKUP(25,A1:C3,3)", "VLOOKUP(5,A1:C3,3)",
    "INDEX(A1:C3,2,3)", "INDEX(C1:C3,3)", "MATCH(\"banana\",C1:C3,0)",
    "MATCH(25,A1:A3,1)", "SUMIFS(A1:A3,C1:C3,\"apple\")",
    "COUNTIFS(C1:C3,\"apple\",A1:A3,\">15\")", "AVERAGEIFS(A1:A3,C1:C3,\"apple\")",
]


def fmt(v):
    if isinstance(v, ExcelError):
        return v.code
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v)


def main():
    ev = Evaluator(lambda sheet, ref: GRID.get(ref),
                   lambda sheet, a, b: [GRID.get(r) for r in iter_range(a, b)],
                   "Sheet1")
    out = [{"formula": f, "python": fmt(ev.eval(parse(f)))} for f in FORMULAS]
    dest = Path(__file__).with_name("parity_expected.json")
    dest.write_text(json.dumps(out, indent=1))
    print(f"wrote {len(out)} cases to {dest}")


if __name__ == "__main__":
    main()
