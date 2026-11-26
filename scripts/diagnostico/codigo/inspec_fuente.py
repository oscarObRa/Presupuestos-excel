import openpyxl
p = r"G:\Mi unidad\Costos\PU2026\OBRA CIVIL, EDIFICACION 2026\CATALOGOS\CATALOGO GENERAL OBRA CIVIL.xlsx"
wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
ws = wb["Sheet"]
rows = list(ws.iter_rows(values_only=True))
print("total filas:", len(rows))
# encontrar fila de encabezado
for i, row in enumerate(rows[:40]):
    txt = " | ".join(str(v) for v in row if v not in (None, ""))
    if txt:
        print(f"F{i+1}: {txt[:300]}")
wb.close()