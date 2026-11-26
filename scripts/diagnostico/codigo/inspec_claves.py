import openpyxl
p = r"G:\Mi unidad\Costos\PU2026\OBRA CIVIL, EDIFICACION 2026\MATRICES\PRELIMINARES\Excavaciones a mano.xlsx"
wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
ws = wb["Sheet"]
filas = list(ws.iter_rows(values_only=True))
clave_rows = [i for i, r in enumerate(filas) if r and r[0] is not None and str(r[0]).strip().lower().startswith("clave")]
print("filas con 'clave:' (1-based):", [c+1 for c in clave_rows][:30])
# mostrar alrededor de la 2da clave
for idx in clave_rows[1:3]:
    print("=" * 60, "alrededor de fila", idx + 1)
    for r in range(max(0, idx - 2), min(len(filas), idx + 50)):
        row = filas[r]
        vals = [(c, str(v)[:18]) for c, v in enumerate(row) if v not in (None, "")]
        if vals:
            print(f" F{r+1}:", vals[:10])
wb.close()