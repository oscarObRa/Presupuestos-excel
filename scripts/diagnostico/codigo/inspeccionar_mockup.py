"""Inspecciona el mockup para replicar el formato de la hoja Presupuesto."""
import openpyxl

ruta = r"H:\Mi unidad\ICEL\PUs\Pretil OMA MTY\Generador_APU_mockup.xlsx"
wb = openpyxl.load_workbook(ruta, data_only=False)

print("=== HOJAS ===")
for i, s in enumerate(wb.sheetnames):
    print(f"  {i}: {s}")

if "Presupuesto" in wb.sheetnames:
    ws = wb["Presupuesto"]
else:
    # buscar hoja que contenga 'presupuesto'
    ws = None
    for s in wb.sheetnames:
        if "presupuesto" in s.lower():
            ws = wb[s]
            print(f"\n(usando hoja: {s})")
            break

if ws is None:
    print("No se encontro hoja Presupuesto")
    raise SystemExit

print(f"\n=== HOJA '{ws.title}' ===")
print(f"  Dimensiones: {ws.dimensions}")
print(f"  Max fila: {ws.max_row}, Max col: {ws.max_column}")
print(f"  Merged cells: {len(ws.merged_cells.ranges)}")

print("\n=== CONTENIDO CELDA POR CELDA (primeras 50 filas) ===")
for r in range(1, min(ws.max_row, 50) + 1):
    fila_datos = []
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=r, column=c)
        v = cell.value
        if v is not None:
            fila_datos.append(f"{cell.coordinate}={repr(v)}")
    if fila_datos:
        print(f"  Fila {r}: " + " | ".join(fila_datos))
