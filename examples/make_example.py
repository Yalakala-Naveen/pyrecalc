"""Generate the example workbook: inputs, formulas, an output sheet."""
import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Inputs"
ws["A1"] = "rate"; ws["B1"] = 0.10
ws["A2"] = "principal"; ws["B2"] = 1000
out = wb.create_sheet("Outputs")
out["A1"] = "=Inputs!B2*(1+Inputs!B1)"
out["A2"] = "=SUM(Inputs!B1:B2)"
out["A3"] = '=IF(A1>1000,"over","under")'
wb.save("examples/model.xlsx")
print("wrote examples/model.xlsx")
