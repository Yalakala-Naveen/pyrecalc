"""pyrecalc — evaluate Excel formulas in Python, without Excel.

The generated-workbook verification wedge: openpyxl writes formulas but
never evaluates them, so pyrecalc computes the values in-process and can
check a workbook's cached values against its own formulas.
"""

from .errors import ExcelError
from .workbook import Mismatch, VerifyReport, Workbook

__version__ = "0.1.0"
__all__ = ["Workbook", "VerifyReport", "Mismatch", "ExcelError", "__version__"]
