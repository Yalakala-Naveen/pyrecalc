"""Workbook: load an .xlsx, evaluate its formulas, verify cached values."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .errors import CIRC, ExcelError
from .evaluator import Evaluator, iter_range
from .parser import normalize_ref, parse

_SHEET_REF = re.compile(r"^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$")


def split_address(address: str) -> tuple[str | None, str]:
    """Split 'Sheet!A1' / "'My Sheet'!A1" / 'A1' into (sheet, ref)."""
    m = _SHEET_REF.match(address.strip())
    if not m:
        return None, normalize_ref(address.strip())
    sheet = (m.group(1) or m.group(2) or "").replace("''", "'")
    return sheet, normalize_ref(m.group(3))


@dataclass
class Mismatch:
    sheet: str
    ref: str
    formula: str
    computed: object
    cached: object


@dataclass
class VerifyReport:
    checked: int = 0
    mismatches: list[Mismatch] = field(default_factory=list)
    errors: list[Mismatch] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.mismatches and not self.errors


def _close(a: object, b: object) -> bool:
    if isinstance(a, ExcelError) or isinstance(b, ExcelError):
        return a == b
    if isinstance(a, float) and isinstance(b, (int, float)):
        return abs(a - float(b)) <= 1e-9 * max(1.0, abs(float(b)))
    return a == b


class Workbook:
    """An .xlsx workbook whose formulas pyrecalc evaluates in-process."""

    def __init__(self, path: str) -> None:
        import openpyxl

        self.path = path
        wb_formulas = openpyxl.load_workbook(path, data_only=False)
        wb_cached = openpyxl.load_workbook(path, data_only=True)
        self._literals: dict[tuple[str, str], object] = {}
        self._formulas: dict[tuple[str, str], str] = {}
        self._cached: dict[tuple[str, str], object] = {}
        self._overrides: dict[tuple[str, str], object] = {}
        self._memo: dict[tuple[str, str], object] = {}
        self._active: set[tuple[str, str]] = set()
        self.sheet_names: list[str] = list(wb_formulas.sheetnames)
        for name in self.sheet_names:
            ws = wb_formulas[name]
            ws_cached = wb_cached[name]
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is None:
                        continue
                    key = (name, cell.coordinate)
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        self._formulas[key] = cell.value[1:]
                        cached = ws_cached[cell.coordinate].value
                        if cached is not None:
                            self._cached[key] = cached
                    else:
                        self._literals[key] = cell.value

    # public API ---------------------------------------------------------
    def set(self, address: str, value: object) -> None:
        """Override a cell's input value and invalidate computed results."""
        sheet, ref = split_address(address)
        key = (sheet or self.sheet_names[0], ref)
        self._overrides[key] = value
        self._memo.clear()

    def value(self, address: str) -> object:
        """The computed value of a cell: 'A1' or 'Sheet!A1'."""
        sheet, ref = split_address(address)
        return self._resolve(sheet or self.sheet_names[0], ref)

    def formula_cells(self) -> list[tuple[str, str, str]]:
        return [(s, r, f) for (s, r), f in sorted(self._formulas.items())]

    def verify(self) -> VerifyReport:
        """Compare every formula cell's computed value to its cached value.

        A generated workbook whose cached values disagree with its own
        formulas is silently wrong — this is the check that catches it.
        Formula cells with no cached value (never opened in Excel) are
        reported as errors only when their computed value is an error;
        otherwise they are counted as checked with nothing to contradict.
        """
        report = VerifyReport()
        for (sheet, ref), formula in sorted(self._formulas.items()):
            computed = self._resolve(sheet, ref)
            report.checked += 1
            entry = Mismatch(sheet, ref, "=" + formula, computed,
                             self._cached.get((sheet, ref)))
            if isinstance(computed, ExcelError):
                report.errors.append(entry)
            elif (sheet, ref) in self._cached and not _close(computed, self._cached[(sheet, ref)]):
                report.mismatches.append(entry)
        return report

    # internals -----------------------------------------------------------
    def _resolve(self, sheet: str, ref: str) -> object:
        key = (sheet, ref)
        if key in self._overrides:
            return self._overrides[key]
        if key in self._memo:
            return self._memo[key]
        if key in self._formulas:
            if key in self._active:
                return CIRC
            self._active.add(key)
            try:
                ev = Evaluator(
                    lambda s, r: self._resolve(s or sheet, r),
                    lambda s, a, b: [self._resolve(s or sheet, cell)
                                      for cell in iter_range(a, b)],
                    sheet,
                )
                result = ev.eval(parse(self._formulas[key]))
            except Exception:  # a broken formula is a #VALUE!, not a crash
                from .errors import VALUE

                result = VALUE
            finally:
                self._active.discard(key)
            self._memo[key] = result
            return result
        return self._literals.get(key)
