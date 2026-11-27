"""Tests for pyrecalc. Fixtures are generated workbooks built with openpyxl,
including planted cached-value mismatches (written by hand-editing the
cached sheet the way a stale Excel save would leave them)."""

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

import openpyxl  # noqa: E402

from pyrecalc import Workbook  # noqa: E402
from pyrecalc.__main__ import main  # noqa: E402
from pyrecalc.errors import ExcelError  # noqa: E402


def build(cells, sheet="Sheet", extra_sheets=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for ref, val in cells.items():
        ws[ref] = val
    for name, scells in (extra_sheets or {}).items():
        ws2 = wb.create_sheet(name)
        for ref, val in scells.items():
            ws2[ref] = val
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    wb.save(path)
    return path


class EvaluatorTests(unittest.TestCase):
    def val(self, formula, cells=None, sheet="Sheet"):
        all_cells = dict(cells or {})
        all_cells["Z9"] = formula
        path = build(all_cells, sheet=sheet)
        self.addCleanup(os.unlink, path)
        return Workbook(path).value("Z9")

    def test_arithmetic_precedence(self):
        self.assertEqual(self.val("=1+2*3"), 7.0)
        self.assertEqual(self.val("=(1+2)*3"), 9.0)
        self.assertEqual(self.val("=2^3"), 8.0)
        self.assertEqual(self.val("=50%"), 0.5)
        self.assertEqual(self.val("=-3+1"), -2.0)

    def test_cell_refs_and_ranges(self):
        cells = {"A1": 10, "A2": 20, "A3": 30}
        self.assertEqual(self.val("=SUM(A1:A3)", cells), 60.0)
        self.assertEqual(self.val("=AVERAGE(A1:A3)", cells), 20.0)
        self.assertEqual(self.val("=$A$1+A2", cells), 30.0)
        self.assertEqual(self.val("=COUNT(A1:A3)", cells), 3.0)
        self.assertEqual(self.val("=MIN(A1:A3)+MAX(A1:A3)", cells), 40.0)

    def test_if_and_logic(self):
        cells = {"A1": 5}
        self.assertEqual(self.val('=IF(A1>3,"big","small")', cells), "big")
        self.assertEqual(self.val('=IF(A1>30,"big","small")', cells), "small")
        self.assertEqual(self.val("=AND(A1>1,A1<10)", cells), True)
        self.assertEqual(self.val("=NOT(A1>10)", cells), True)

    def test_text_functions(self):
        self.assertEqual(self.val('="a"&"b"'), "ab")
        self.assertEqual(self.val('=CONCAT("x","y")'), "xy")
        self.assertEqual(self.val('=LEN("hello")'), 5.0)
        self.assertEqual(self.val('=UPPER("ab")'), "AB")

    def test_math_functions(self):
        self.assertEqual(self.val("=ROUND(2.5,0)"), 3.0)
        self.assertEqual(self.val("=ROUND(2.567,2)"), 2.57)
        self.assertEqual(self.val("=ABS(-4)"), 4.0)
        self.assertEqual(self.val("=SQRT(9)"), 3.0)
        self.assertEqual(self.val("=POWER(2,10)"), 1024.0)

    def test_criteria_functions(self):
        cells = {"A1": 10, "A2": 20, "A3": 30, "B1": 1, "B2": 2, "B3": 3}
        self.assertEqual(self.val("=SUMIF(A1:A3,\">15\")", cells), 50.0)
        self.assertEqual(self.val("=SUMIF(A1:A3,\">15\",B1:B3)", cells), 5.0)
        self.assertEqual(self.val("=COUNTIF(A1:A3,\">=20\")", cells), 2.0)
        self.assertEqual(self.val("=AVERAGEIF(A1:A3,\"<25\")", cells), 15.0)
        self.assertEqual(self.val("=AVERAGEIF(A1:A3,\">100\")", cells),
                         ExcelError("#DIV/0!"))
        labels = {"A1": "Apple", "A2": "apple", "A3": "Pear", "B1": 4, "B2": 6, "B3": 9}
        self.assertEqual(self.val('=SUMIF(A1:A3,"apple",B1:B3)', labels), 10.0)
        self.assertEqual(self.val('=COUNTIF(A1:A3,"<>apple")', labels), 1.0)

    def test_lookup_functions(self):
        # price table: code, name, price in A1:C3
        cells = {"A1": "a1", "B1": "Apple", "C1": 4,
                 "A2": "b2", "B2": "Pear", "C2": 6,
                 "A3": "c3", "B3": "Fig", "C3": 9}
        self.assertEqual(self.val('=VLOOKUP("b2",A1:C3,3,FALSE)', cells), 6.0)
        self.assertEqual(self.val('=VLOOKUP("B2",A1:C3,2,FALSE)', cells), "Pear")
        self.assertEqual(self.val('=VLOOKUP("zz",A1:C3,2,FALSE)', cells),
                         ExcelError("#N/A"))
        self.assertEqual(self.val('=VLOOKUP("a1",A1:C3,4,FALSE)', cells),
                         ExcelError("#REF!"))
        self.assertEqual(self.val('=INDEX(A1:C3,2,3)', cells), 6.0)
        self.assertEqual(self.val('=INDEX(C1:C3,3)', cells), 9.0)
        self.assertEqual(self.val('=INDEX(A1:C3,9,1)', cells),
                         ExcelError("#REF!"))
        self.assertEqual(self.val('=MATCH("c3",A1:A3,0)', cells), 3.0)
        self.assertEqual(self.val('=MATCH("zz",A1:A3,0)', cells),
                         ExcelError("#N/A"))
        # lookups compose: INDEX+MATCH is the standard VLOOKUP replacement
        self.assertEqual(self.val('=INDEX(B1:B3,MATCH("a1",A1:A3,0))', cells),
                         "Apple")

    def test_approximate_lookups(self):
        # sorted tier table: threshold, label
        cells = {"A1": 0, "B1": "free", "A2": 100, "B2": "pro",
                 "A3": 500, "B3": "team"}
        self.assertEqual(self.val('=VLOOKUP(250,A1:B3,2)', cells), "pro")
        self.assertEqual(self.val('=VLOOKUP(500,A1:B3,2,TRUE)', cells), "team")
        self.assertEqual(self.val('=VLOOKUP(-1,A1:B3,2)', cells),
                         ExcelError("#N/A"))
        self.assertEqual(self.val('=MATCH(250,A1:A3,1)', cells), 2.0)
        self.assertEqual(self.val('=MATCH(100,A1:A3,1)', cells), 2.0)

    def test_multi_criteria_functions(self):
        cells = {"A1": "East", "B1": "Apple", "C1": 10,
                 "A2": "East", "B2": "Pear", "C2": 20,
                 "A3": "West", "B3": "Apple", "C3": 30}
        self.assertEqual(
            self.val('=SUMIFS(C1:C3,A1:A3,"East",B1:B3,"Apple")', cells), 10.0)
        self.assertEqual(
            self.val('=COUNTIFS(A1:A3,"East",B1:B3,"<>Fig")', cells), 2.0)
        self.assertEqual(
            self.val('=AVERAGEIFS(C1:C3,B1:B3,"Apple")', cells), 20.0)
        self.assertEqual(
            self.val('=AVERAGEIFS(C1:C3,B1:B3,"Fig")', cells),
            ExcelError("#DIV/0!"))
        self.assertEqual(
            self.val('=SUMIFS(C1:C3,A1:A3,">=100")', cells), 0.0)

    def test_sumproduct_and_rounding_family(self):
        cells = {"A1": 2, "A2": 3, "B1": 10, "B2": 20}
        self.assertEqual(self.val("=SUMPRODUCT(A1:A2,B1:B2)", cells), 80.0)
        self.assertEqual(self.val("=ROUNDUP(2.51,0)"), 3.0)
        self.assertEqual(self.val("=ROUNDDOWN(2.99,0)"), 2.0)
        self.assertEqual(self.val("=ROUNDUP(-2.51,0)"), -3.0)
        self.assertEqual(self.val("=INT(-1.5)"), -2.0)
        self.assertEqual(self.val("=MOD(7,3)"), 1.0)
        self.assertEqual(self.val("=MOD(7,0)"), ExcelError("#DIV/0!"))

    def test_more_text_functions(self):
        self.assertEqual(self.val('=LEFT("hello",2)'), "he")
        self.assertEqual(self.val('=RIGHT("hello",2)'), "lo")
        self.assertEqual(self.val('=MID("hello",2,3)'), "ell")
        self.assertEqual(self.val('=TRIM("  a   b ")'), "a b")
        self.assertEqual(self.val('=PROPER("jane doe")'), "Jane Doe")
        self.assertEqual(self.val('=LOWER("AB")'), "ab")

    def test_errors_are_values(self):
        self.assertEqual(self.val("=1/0"), ExcelError("#DIV/0!"))
        self.assertEqual(self.val("=NOSUCHFN(1)"), ExcelError("#NAME?"))
        self.assertEqual(self.val('=IFERROR(1/0,"fallback")'), "fallback")
        # error propagates through a dependent formula
        path = build({"A1": "=1/0", "A2": "=A1+1"})
        self.addCleanup(os.unlink, path)
        self.assertEqual(Workbook(path).value("A2"), ExcelError("#DIV/0!"))

    def test_comparisons(self):
        self.assertEqual(self.val("=1<2"), True)
        self.assertEqual(self.val("=2<>2"), False)
        self.assertEqual(self.val('="a"="A"'), True)  # Excel compares case-insensitively

    def test_cross_sheet_ref(self):
        path = build({"A1": "=Inputs!B3*2"}, extra_sheets={"Inputs": {"B3": 21}})
        self.addCleanup(os.unlink, path)
        wb = Workbook(path)
        self.assertEqual(wb.value("A1"), 42.0)
        wb.set("Inputs!B3", 0.12)
        self.assertEqual(wb.value("A1"), 0.24)

    def test_circular_reference_detected(self):
        path = build({"A1": "=A2+1", "A2": "=A1+1"})
        self.addCleanup(os.unlink, path)
        self.assertEqual(Workbook(path).value("A1"), ExcelError("#CIRC!"))


class VerifyTests(unittest.TestCase):
    def test_verify_clean_and_mismatch(self):
        # openpyxl writes no cached values, so verify() has nothing to
        # contradict: checked counts, no mismatches.
        path = build({"A1": 2, "A2": "=A1*3"})
        self.addCleanup(os.unlink, path)
        report = Workbook(path).verify()
        self.assertEqual(report.checked, 1)
        self.assertTrue(report.ok)

        # a stale cached value (what an out-of-date Excel save leaves
        # behind) that disagrees with the formula must be a mismatch.
        wb = Workbook(path)
        wb._cached[("Sheet", "A2")] = 999
        report = wb.verify()
        self.assertFalse(report.ok)
        self.assertEqual(len(report.mismatches), 1)
        self.assertEqual(report.mismatches[0].computed, 6.0)

    def test_cli_value_and_check(self):
        path = build({"A1": 3, "A2": "=A1*A1"})
        self.addCleanup(os.unlink, path)
        self.assertEqual(main(["value", path, "A2"]), 0)
        self.assertEqual(main(["check", path]), 0)
        self.assertEqual(main(["bogus", path]), 2)


if __name__ == "__main__":
    unittest.main()
