"""CLI: pyrecalc check FILE | pyrecalc value FILE CELL.

Exit contract: 0 = ok, 1 = mismatch / error value found, 2 = tool error.
"""

from __future__ import annotations

import sys

from .workbook import Workbook


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if len(args) < 2 or args[0] not in ("check", "value"):
        print("usage: pyrecalc check FILE | pyrecalc value FILE CELL", file=sys.stderr)
        return 2
    try:
        wb = Workbook(args[1])
    except Exception as exc:
        print(f"error: cannot open {args[1]}: {exc}", file=sys.stderr)
        return 2
    if args[0] == "value":
        if len(args) != 3:
            print("usage: pyrecalc value FILE CELL", file=sys.stderr)
            return 2
        print(wb.value(args[2]))
        return 0
    report = wb.verify()
    for m in report.mismatches:
        print(f"MISMATCH {m.sheet}!{m.ref} {m.formula} computed={m.computed!r} cached={m.cached!r}")
    for m in report.errors:
        print(f"ERROR    {m.sheet}!{m.ref} {m.formula} computed={m.computed!r}")
    print(f"{report.checked} formula cells checked, "
          f"{len(report.mismatches)} mismatches, {len(report.errors)} errors")
    return 0 if report.ok else 1


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
