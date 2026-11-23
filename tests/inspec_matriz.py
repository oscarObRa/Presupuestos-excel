import openpyxl
p = r"G:\Mi unidad\Costos\PU2026\OBRA CIVIL, EDIFICACION 2026\MATRICES\PRELIMINARES\Excavaciones a mano.xlsx"
wb = openpyxl.load_workbook(p, data_only=True)
print("HOJAS:", wb.sheetnames)
for name in wb.sheetnames[:3]:
    ws = wb[name]
    print("=" * 70)
    print("HOJA:", name, ws.max_row, "x", ws.max_column)
    for r in range(1, min(ws.max_row, 45) + 1):
        vals = []
        for c in range(1, min(ws.max_column, 8) + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None:
                vals.append(f"{ws.cell(row=r,column=c).coordinate}={str(v)[:26]}")
        if vals:
            print(" F", r, " | ".join(vals))
wb.close()