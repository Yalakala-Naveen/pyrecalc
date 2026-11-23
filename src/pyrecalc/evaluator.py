"""Evaluate parsed formulas against a cell-value resolver.

The resolver is supplied by the Workbook: ``resolve(sheet, ref)`` returns a
cell's value (a literal, or the recursively evaluated formula result) and
``resolve_range(sheet, start, end)`` returns a flat list of values.
"""

from __future__ import annotations

import math
import re
from typing import Callable, Iterable

from .errors import CIRC, DIV0, NA, NAME, REF, VALUE, ExcelError

Resolver = Callable[[str | None, str], object]
RangeResolver = Callable[[str | None, str, str], list]

_CELL_SPLIT = re.compile(r"^([A-Z]+)(\d+)$")


def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def split_ref(ref: str) -> tuple[int, int]:
    m = _CELL_SPLIT.match(ref)
    assert m is not None
    return col_to_num(m.group(1)), int(m.group(2))


def num_to_col(n: int) -> str:
    out = ""
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def iter_range(start: str, end: str) -> Iterable[str]:
    c1, r1 = split_ref(start)
    c2, r2 = split_ref(end)
    for r in range(min(r1, r2), max(r1, r2) + 1):
        for c in range(min(c1, c2), max(c1, c2) + 1):
            yield f"{num_to_col(c)}{r}"


def _is_num(v: object) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _to_num(v: object) -> object:
    """Coerce a scalar the way Excel arithmetic does, or an ExcelError."""
    if isinstance(v, ExcelError):
        return v
    if v is None:
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if _is_num(v):
        return float(v)  # type: ignore[arg-type]
    if isinstance(v, str):
        try:
            return float(v.strip())
        except ValueError:
            return VALUE
    return VALUE


def _to_text(v: object) -> object:
    if isinstance(v, ExcelError):
        return v
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if _is_num(v):
        f = float(v)  # type: ignore[arg-type]
        return str(int(f)) if f == int(f) else repr(f)
    return str(v)


def _flatten(args: list) -> list:
    out: list = []
    for a in args:
        if isinstance(a, list):
            out.extend(a)
        else:
            out.append(a)
    return out


class Evaluator:
    def __init__(self, resolve: Resolver, resolve_range: RangeResolver,
                 current_sheet: str) -> None:
        self.resolve = resolve
        self.resolve_range = resolve_range
        self.current_sheet = current_sheet

    def eval(self, node) -> object:
        kind = node[0]
        if kind in ("num", "str", "bool"):
            return node[1]
        if kind == "name":
            return NAME
        if kind == "ref":
            sheet = node[1] or self.current_sheet
            return self.resolve(sheet, node[2])
        if kind == "range":
            sheet = node[1] or self.current_sheet
            return self.resolve_range(sheet, node[2], node[3])
        if kind == "percent":
            v = _to_num(self.eval(node[1]))
            return v if isinstance(v, ExcelError) else v / 100.0
        if kind == "unary":
            v = _to_num(self.eval(node[2]))
            if isinstance(v, ExcelError):
                return v
            return -v if node[1] == "-" else v
        if kind == "binop":
            return self._binop(node[1], node[2], node[3])
        if kind == "call":
            return self._call(node[1], node[2])
        raise ValueError(f"unknown AST node {kind!r}")  # pragma: no cover

    def _binop(self, op: str, lnode, rnode) -> object:
        if op == "&":
            lt, rt = _to_text(self.eval(lnode)), _to_text(self.eval(rnode))
            if isinstance(lt, ExcelError):
                return lt
            if isinstance(rt, ExcelError):
                return rt
            return str(lt) + str(rt)
        left, right = self.eval(lnode), self.eval(rnode)
        if isinstance(left, list) or isinstance(right, list):
            return VALUE
        if op in ("=", "<>", "<", "<=", ">", ">="):
            return self._compare(op, left, right)
        ln, rn = _to_num(left), _to_num(right)
        if isinstance(ln, ExcelError):
            return ln
        if isinstance(rn, ExcelError):
            return rn
        if op == "+":
            return ln + rn  # type: ignore[operator]
        if op == "-":
            return ln - rn  # type: ignore[operator]
        if op == "*":
            return ln * rn  # type: ignore[operator]
        if op == "/":
            return DIV0 if rn == 0 else ln / rn  # type: ignore[operator]
        if op == "^":
            try:
                return math.pow(ln, rn)  # type: ignore[arg-type]
            except (OverflowError, ValueError):
                return VALUE
        return NAME  # pragma: no cover

    @staticmethod
    def _compare(op: str, left: object, right: object) -> object:
        if isinstance(left, ExcelError):
            return left
        if isinstance(right, ExcelError):
            return right
        if _is_num(left) and _is_num(right):
            a, b = float(left), float(right)  # type: ignore[arg-type]
        elif isinstance(left, str) and isinstance(right, str):
            a, b = left.casefold(), right.casefold()
        elif isinstance(left, bool) and isinstance(right, bool):
            a, b = left, right
        elif left is None and _is_num(right):
            a, b = 0.0, float(right)  # type: ignore[arg-type]
        elif right is None and _is_num(left):
            a, b = float(left), 0.0  # type: ignore[arg-type]
        else:
            return VALUE
        return {"=": a == b, "<>": a != b, "<": a < b,
                "<=": a <= b, ">": a > b, ">=": a >= b}[op]

    # functions ---------------------------------------------------------
    def _call(self, name: str, arg_nodes: list) -> object:
        if name == "IF":
            if len(arg_nodes) not in (2, 3):
                return VALUE
            cond = self.eval(arg_nodes[0])
            if isinstance(cond, ExcelError):
                return cond
            truthy = cond if isinstance(cond, bool) else (_to_num(cond) not in (0.0,) if _is_num(_to_num(cond)) else False)
            if truthy:
                return self.eval(arg_nodes[1])
            return self.eval(arg_nodes[2]) if len(arg_nodes) == 3 else False
        if name == "IFERROR":
            if len(arg_nodes) != 2:
                return VALUE
            v = self.eval(arg_nodes[0])
            return self.eval(arg_nodes[1]) if isinstance(v, ExcelError) else v
        if name in ("AND", "OR", "NOT"):
            vals = [self.eval(a) for a in arg_nodes]
            for v in _flatten(vals):
                if isinstance(v, ExcelError):
                    return v
            flat = [v for v in _flatten(vals) if v is not None]
            bools = [v if isinstance(v, bool) else (float(v) != 0.0 if _is_num(v) else None) for v in flat]
            if any(b is None for b in bools):
                return VALUE
            if name == "AND":
                return all(bools)
            if name == "OR":
                return any(bools)
            return not bools[0] if len(bools) == 1 else VALUE

        args = [self.eval(a) for a in arg_nodes]
        flat = _flatten(args)
        for v in flat:
            if isinstance(v, ExcelError):
                return v
        nums = [float(v) for v in flat if _is_num(v)]

        if name == "SUM":
            return sum(nums)
        if name == "AVERAGE":
            return sum(nums) / len(nums) if nums else DIV0
        if name == "MIN":
            return min(nums) if nums else 0.0
        if name == "MAX":
            return max(nums) if nums else 0.0
        if name == "COUNT":
            return float(len(nums))
        if name == "COUNTA":
            return float(len([v for v in flat if v is not None and v != ""]))
        if name == "ROUND":
            if len(args) != 2 or not _is_num(args[0]) or not _is_num(args[1]):
                return VALUE
            places = int(args[1])  # type: ignore[arg-type]
            factor = 10.0 ** places
            return math.floor(abs(float(args[0])) * factor + 0.5) / factor * (1 if float(args[0]) >= 0 else -1)  # type: ignore[arg-type]
        if name == "ABS":
            return abs(float(args[0])) if len(args) == 1 and _is_num(args[0]) else VALUE
        if name == "SQRT":
            if len(args) == 1 and _is_num(args[0]):
                return math.sqrt(float(args[0])) if float(args[0]) >= 0 else VALUE  # type: ignore[arg-type]
            return VALUE
        if name == "POWER":
            if len(args) == 2 and _is_num(args[0]) and _is_num(args[1]):
                try:
                    return math.pow(float(args[0]), float(args[1]))  # type: ignore[arg-type]
                except (OverflowError, ValueError):
                    return VALUE
            return VALUE
        if name in ("CONCAT", "CONCATENATE"):
            parts = [_to_text(v) for v in flat]
            if any(isinstance(p, ExcelError) for p in parts):
                return VALUE
            return "".join(str(p) for p in parts)
        if name == "LEN":
            t = _to_text(args[0]) if len(args) == 1 else VALUE
            return VALUE if isinstance(t, ExcelError) else float(len(str(t)))
        if name in ("UPPER", "LOWER"):
            t = _to_text(args[0]) if len(args) == 1 else VALUE
            if isinstance(t, ExcelError):
                return t
            return str(t).upper() if name == "UPPER" else str(t).lower()
        if name == "NA":
            return NA
        return NAME
