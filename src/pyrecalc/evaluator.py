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


def _criteria_match(value: object, criteria: object) -> bool:
    """Excel criteria semantics: an optional comparison operator prefix
    (``>=10``, ``<>x``, ``=foo``) or a plain value compared for equality
    (numbers numerically, text case-insensitively)."""
    if isinstance(criteria, ExcelError):
        return False
    op, target = "=", criteria
    if isinstance(criteria, str):
        for prefix in ("<>", "<=", ">=", "<", ">", "="):
            if criteria.startswith(prefix):
                op, target = prefix, criteria[len(prefix):]
                break
        if isinstance(target, str):
            try:
                target = float(target)
            except ValueError:
                pass
    if _is_num(value) and _is_num(target):
        a, b = float(value), float(target)  # type: ignore[arg-type]
        return {"=": a == b, "<>": a != b, "<": a < b,
                "<=": a <= b, ">": a > b, ">=": a >= b}[op]
    if op in ("=", "<>"):
        eq = (value is None and target in ("", None)) or (
            isinstance(value, str) and isinstance(target, str)
            and value.casefold() == target.casefold()) or value == target
        return eq if op == "=" else not eq
    return False


def _range_width(node) -> int | None:
    """Column count of a range/ref AST node, or None for anything else."""
    if node[0] == "range":
        c1, _ = split_ref(node[2])
        c2, _ = split_ref(node[3])
        return abs(c2 - c1) + 1
    if node[0] == "ref":
        return 1
    return None


def _lookup_cmp(value: object, target: object) -> int | None:
    """-1/0/1 comparison for approximate lookups, or None if the pair is
    not comparable (mixed types never match in an approximate scan)."""
    if _is_num(value) and _is_num(target):
        a, b = float(value), float(target)  # type: ignore[arg-type]
    elif isinstance(value, str) and isinstance(target, str):
        a, b = value.casefold(), target.casefold()
    else:
        return None
    return (a > b) - (a < b)  # type: ignore[operator]


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
        if name in ("SUMIFS", "COUNTIFS", "AVERAGEIFS"):
            # Multi-criteria aggregation. SUMIFS/AVERAGEIFS take the
            # value range FIRST, then (criteria_range, criteria) pairs;
            # COUNTIFS is pairs only. Ranges stay separate, as in SUMIF.
            lead = 0 if name == "COUNTIFS" else 1
            if len(arg_nodes) < lead + 2 or (len(arg_nodes) - lead) % 2:
                return VALUE
            ranges: list[list] = []
            criteria: list[object] = []
            for i in range(lead, len(arg_nodes), 2):
                rv = self.eval(arg_nodes[i])
                if isinstance(rv, ExcelError):
                    return rv
                ranges.append(rv if isinstance(rv, list) else [rv])
                cv = self.eval(arg_nodes[i + 1])
                if isinstance(cv, ExcelError):
                    return cv
                criteria.append(cv)
            if len({len(r) for r in ranges}) != 1:
                return VALUE
            for r in ranges:
                for v in r:
                    if isinstance(v, ExcelError):
                        return v
            keep = [i for i in range(len(ranges[0]))
                    if all(_criteria_match(r[i], c)
                           for r, c in zip(ranges, criteria))]
            if name == "COUNTIFS":
                return float(len(keep))
            vals = self.eval(arg_nodes[0])
            if isinstance(vals, ExcelError):
                return vals
            sum_vals = vals if isinstance(vals, list) else [vals]
            if len(sum_vals) != len(ranges[0]):
                return VALUE
            matched = [float(sum_vals[i]) for i in keep
                       if _is_num(sum_vals[i])]
            if name == "SUMIFS":
                return sum(matched)
            return sum(matched) / len(matched) if matched else DIV0
        if name == "XLOOKUP":
            # XLOOKUP(lookup, lookup_array, return_array,
            #          [if_not_found], [match_mode], [search_mode])
            # match_mode: 0 exact (default), -1 exact-or-next-smaller,
            # 1 exact-or-next-larger. search_mode: 1 first-to-last
            # (default), -1 last-to-first; binary modes (2/-2) fall back
            # to a linear scan, which returns the same value on the
            # sorted arrays binary search requires.
            if len(arg_nodes) not in (3, 4, 5, 6):
                return VALUE
            lookup = self.eval(arg_nodes[0])
            if isinstance(lookup, ExcelError):
                return lookup
            la = self.eval(arg_nodes[1])
            if isinstance(la, ExcelError):
                return la
            ra = self.eval(arg_nodes[2])
            if isinstance(ra, ExcelError):
                return ra
            keys = la if isinstance(la, list) else [la]
            rets = ra if isinstance(ra, list) else [ra]
            if not keys or len(rets) % len(keys):
                return VALUE
            per = len(rets) // len(keys)  # return columns per key
            match_mode = 0.0
            if len(arg_nodes) >= 5:
                mm = self.eval(arg_nodes[4])
                if isinstance(mm, ExcelError):
                    return mm
                if not _is_num(mm):
                    return VALUE
                match_mode = float(mm)  # type: ignore[arg-type]
            search_mode = 1.0
            if len(arg_nodes) >= 6:
                sm = self.eval(arg_nodes[5])
                if isinstance(sm, ExcelError):
                    return sm
                if not _is_num(sm):
                    return VALUE
                search_mode = float(sm)  # type: ignore[arg-type]
            order = range(len(keys) - 1, -1, -1) if search_mode < 0 else range(len(keys))
            found: int | None = None
            for i in order:
                if _criteria_match(keys[i], lookup):
                    found = i
                    break
            if found is None and match_mode in (-1.0, 1.0):
                for i in order:
                    cmp = _lookup_cmp(keys[i], lookup)
                    if cmp is None or cmp == 0:
                        continue
                    if match_mode == -1.0 and cmp < 0:
                        if found is None or _lookup_cmp(keys[i], keys[found]) == 1:  # type: ignore[arg-type]
                            found = i
                    if match_mode == 1.0 and cmp > 0:
                        if found is None or _lookup_cmp(keys[i], keys[found]) == -1:  # type: ignore[arg-type]
                            found = i
            if found is None:
                if len(arg_nodes) >= 4:
                    return self.eval(arg_nodes[3])
                return NA
            row = rets[found * per:(found + 1) * per]
            return row[0] if per == 1 else row
        if name == "VLOOKUP":
            if len(arg_nodes) not in (3, 4):
                return VALUE
            lookup = self.eval(arg_nodes[0])
            if isinstance(lookup, ExcelError):
                return lookup
            table = self.eval(arg_nodes[1])
            if isinstance(table, ExcelError):
                return table
            width = _range_width(arg_nodes[1])
            if width is None:
                return VALUE
            cells = table if isinstance(table, list) else [table]
            rows = [cells[i:i + width] for i in range(0, len(cells), width)]
            col = self.eval(arg_nodes[2])
            if not _is_num(col):
                return VALUE
            idx = int(col)  # type: ignore[arg-type]
            if idx < 1 or idx > width:
                return REF
            exact = False
            if len(arg_nodes) == 4:
                flag = self.eval(arg_nodes[3])
                if isinstance(flag, ExcelError):
                    return flag
                exact = flag is False or (_is_num(flag) and float(flag) == 0.0)  # type: ignore[arg-type]
            if exact:
                for row in rows:
                    if _criteria_match(row[0], lookup):
                        return row[idx - 1]
                return NA
            found = None
            for row in rows:
                cmp = _lookup_cmp(row[0], lookup)
                if cmp is not None and cmp <= 0:
                    found = row
            return found[idx - 1] if found is not None else NA
        if name == "INDEX":
            if len(arg_nodes) not in (2, 3):
                return VALUE
            width = _range_width(arg_nodes[0])
            if width is None:
                return VALUE
            arr = self.eval(arg_nodes[0])
            if isinstance(arr, ExcelError):
                return arr
            cells = arr if isinstance(arr, list) else [arr]
            rows = [cells[i:i + width] for i in range(0, len(cells), width)]
            rn = self.eval(arg_nodes[1])
            if not _is_num(rn):
                return VALUE
            r = int(rn)  # type: ignore[arg-type]
            if len(arg_nodes) == 2 and (width == 1 or len(rows) == 1):
                flat = cells
                return flat[r - 1] if 1 <= r <= len(flat) else REF
            if r < 1 or r > len(rows):
                return REF
            if len(arg_nodes) == 2:
                return rows[r - 1]
            cn = self.eval(arg_nodes[2])
            if not _is_num(cn):
                return VALUE
            c = int(cn)  # type: ignore[arg-type]
            return rows[r - 1][c - 1] if 1 <= c <= width else REF
        if name == "MATCH":
            if len(arg_nodes) not in (2, 3):
                return VALUE
            lookup = self.eval(arg_nodes[0])
            if isinstance(lookup, ExcelError):
                return lookup
            arr = self.eval(arg_nodes[1])
            if isinstance(arr, ExcelError):
                return arr
            vals = arr if isinstance(arr, list) else [arr]
            mtype = 1.0
            if len(arg_nodes) == 3:
                t = self.eval(arg_nodes[2])
                if isinstance(t, ExcelError):
                    return t
                if not _is_num(t):
                    return VALUE
                mtype = float(t)  # type: ignore[arg-type]
            if mtype == 0:
                for i, v in enumerate(vals):
                    if _criteria_match(v, lookup):
                        return float(i + 1)
                return NA
            best: int | None = None
            for i, v in enumerate(vals):
                cmp = _lookup_cmp(v, lookup)
                if cmp is None:
                    continue
                if mtype > 0 and cmp <= 0:
                    best = i
                if mtype < 0 and cmp >= 0:
                    best = i
            return float(best + 1) if best is not None else NA
        if name in ("SUMIF", "COUNTIF", "AVERAGEIF"):
            # Criteria functions keep their ranges separate — the generic
            # path below flattens every argument together and would lose
            # which values the criteria applies to.
            if len(arg_nodes) not in (2, 3):
                return VALUE
            crit_range = self.eval(arg_nodes[0])
            criteria = self.eval(arg_nodes[1])
            if isinstance(criteria, ExcelError):
                return criteria
            crit_vals = crit_range if isinstance(crit_range, list) else [crit_range]
            for v in crit_vals:
                if isinstance(v, ExcelError):
                    return v
            if name == "COUNTIF":
                if len(arg_nodes) != 2:
                    return VALUE
                return float(sum(1 for v in crit_vals
                                 if _criteria_match(v, criteria)))
            sum_range = self.eval(arg_nodes[2]) if len(arg_nodes) == 3 else crit_range
            if isinstance(sum_range, ExcelError):
                return sum_range
            sum_vals = sum_range if isinstance(sum_range, list) else [sum_range]
            for v in sum_vals:
                if isinstance(v, ExcelError):
                    return v
            matched = [float(sv) for cv, sv in zip(crit_vals, sum_vals)
                       if _criteria_match(cv, criteria) and _is_num(sv)]
            if name == "SUMIF":
                return sum(matched)
            return sum(matched) / len(matched) if matched else DIV0
        if name == "SUMPRODUCT":
            arrays = []
            for a in arg_nodes:
                v = self.eval(a)
                vals = v if isinstance(v, list) else [v]
                for item in vals:
                    if isinstance(item, ExcelError):
                        return item
                arrays.append([float(x) if _is_num(x) else 0.0 for x in vals])
            if not arrays or len({len(a) for a in arrays}) != 1:
                return VALUE
            return sum(math.prod(col) for col in zip(*arrays))
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
        if name in ("ROUNDUP", "ROUNDDOWN"):
            if len(args) != 2 or not _is_num(args[0]) or not _is_num(args[1]):
                return VALUE
            places = int(args[1])  # type: ignore[arg-type]
            factor = 10.0 ** places
            scaled = abs(float(args[0])) * factor  # type: ignore[arg-type]
            done = math.ceil(scaled) if name == "ROUNDUP" else math.floor(scaled)
            return done / factor * (1 if float(args[0]) >= 0 else -1)  # type: ignore[arg-type]
        if name == "INT":
            return float(math.floor(float(args[0]))) if len(args) == 1 and _is_num(args[0]) else VALUE  # type: ignore[arg-type]
        if name == "MOD":
            if len(args) == 2 and _is_num(args[0]) and _is_num(args[1]):
                if float(args[1]) == 0:  # type: ignore[arg-type]
                    return DIV0
                return float(args[0]) % float(args[1])  # type: ignore[operator]
            return VALUE
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
        if name in ("LEFT", "RIGHT"):
            if len(args) not in (1, 2):
                return VALUE
            t = _to_text(args[0])
            if isinstance(t, ExcelError):
                return t
            n = int(args[1]) if len(args) == 2 and _is_num(args[1]) else 1
            if n < 0:
                return VALUE
            s = str(t)
            return s[:n] if name == "LEFT" else (s[-n:] if n else "")
        if name == "MID":
            if len(args) != 3 or not _is_num(args[1]) or not _is_num(args[2]):
                return VALUE
            t = _to_text(args[0])
            if isinstance(t, ExcelError):
                return t
            start, count = int(args[1]), int(args[2])  # type: ignore[arg-type]
            if start < 1 or count < 0:
                return VALUE
            return str(t)[start - 1:start - 1 + count]
        if name == "TRIM":
            t = _to_text(args[0]) if len(args) == 1 else VALUE
            if isinstance(t, ExcelError):
                return t
            return " ".join(str(t).split())
        if name == "PROPER":
            t = _to_text(args[0]) if len(args) == 1 else VALUE
            if isinstance(t, ExcelError):
                return t
            return str(t).title()
        if name == "NA":
            return NA
        return NAME
