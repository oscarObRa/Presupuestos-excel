import openpyxl
p = r"G:\Mi unidad\Costos\PU2026\OBRA CIVIL, EDIFICACION 2026\MATRICES\PRELIMINARES\Excavaciones a mano.xlsx"
wb = openpyxl.load_workbook(p, data_only=True)
ws = wb["Sheet"]
for r in range(23, 60):
    vals = []
    for c in range(1, 27):
        v = ws.cell(row=r, column=c).value
        if v is not None:
            vals.append(f"{ws.cell(row=r,column=c).coordinate}={str(v)[:24]}")
    if vals:
        print(f"F{r}: " + " | ".join(vals))
wb.close()