"""Excel error values, represented as values (never raised across the API)."""


class ExcelError:
    """An Excel error value such as #DIV/0! or #VALUE!.

    Errors are *values* in Excel: a formula that divides by zero does not
    crash the workbook, it produces #DIV/0! which then propagates into
    every formula that depends on it. pyrecalc mirrors that: evaluation
    returns ExcelError instances instead of raising.
    """

    __slots__ = ("code",)

    def __init__(self, code: str) -> None:
        self.code = code

    def __repr__(self) -> str:  # pragma: no cover - trivial
        return self.code

    def __eq__(self, other: object) -> bool:
        return isinstance(other, ExcelError) and other.code == self.code

    def __hash__(self) -> int:
        return hash((ExcelError, self.code))


DIV0 = ExcelError("#DIV/0!")
VALUE = ExcelError("#VALUE!")
REF = ExcelError("#REF!")
NAME = ExcelError("#NAME?")
NA = ExcelError("#N/A")
CIRC = ExcelError("#CIRC!")


class FormulaSyntaxError(ValueError):
    """Raised by the parser for a formula that cannot be tokenized/parsed."""
