"""Inspección detallada del mockup: estilos, encabezado, columnas, propuesta."""
import openpyxl
from openpyxl.utils import get_column_letter

ruta = r"H:\Mi unidad\ICEL\PUs\Pretil OMA MTY\Generador_APU_mockup.xlsx"
wb = openpyxl.load_workbook(ruta, data_only=False)

def describir_celda(cell):
    f = cell.font
    fill = cell.fill
    al = cell.alignment
    b = cell.border
    parts = []
    if cell.value is not None:
        parts.append(f"val={repr(cell.value)}")
    if f:
        fc = f.color.rgb if f.color and f.color.rgb else None
        parts.append(f"font(bold={f.bold},sz={f.size},color={fc},name={f.name})")
    if fill and fill.fgColor and fill.fgColor.rgb and fill.fgColor.rgb != "00000000":
        parts.append(f"fill={fill.fgColor.rgb}")
    if al and (al.horizontal or al.vertical or al.wrap_text):
        parts.append(f"align(h={al.horizontal},v={al.vertical},wrap={al.wrap_text})")
    if cell.number_format and cell.number_format != "General":
        parts.append(f"fmt={cell.number_format}")
    return " ".join(parts)

# ---- Hoja PROPUESTA ----
print("=" * 70)
print("HOJA PROPUESTA")
print("=" * 70)
ws = wb["PROPUESTA"]
print(f"Dimensiones: {ws.dimensions}, Merged: {len(ws.merged_cells.ranges)}")
for r in range(1, min(ws.max_row, 40) + 1):
    for c in range(1, min(ws.max_column, 12) + 1):
        cell = ws.cell(row=r, column=c)
        if cell.value is not None:
            print(f"  {cell.coordinate}: {describir_celda(cell)}")

# ---- Hoja PResupuesto: encabezado y estilos ----
print()
print("=" * 70)
print("HOJA PResupuesto - ENCABEZADO Y ESTILOS (filas 1-10)")
print("=" * 70)
ws = wb["PResupuesto"]
print(f"Dimensiones: {ws.dimensions}")
print(f"Anchos de columna:")
for col, dim in ws.column_dimensions.items():
    if dim.width:
        print(f"  {col}: {dim.width}")
print(f"\nMerged: {list(ws.merged_cells.ranges)}")
print(f"\nFilas 1-10 (todas las columnas con contenido/estilo):")
for r in range(1, 11):
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=r, column=c)
        d = describir_celda(cell)
        if d:
            print(f"  {cell.coordinate}: {d}")

# Fila de totales (últimas filas)
print(f"\nÚltimas filas (48-52):")
for r in range(48, ws.max_row + 1):
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=r, column=c)
        d = describir_celda(cell)
        if d:
            print(f"  {cell.coordinate}: {d}")

# Config de impresión
print()
print("=" * 70)
print("CONFIG IMPRESIÓN PResupuesto")
print("=" * 70)
ps = ws.page_setup
print(f"  orientation: {ps.orientation}")
print(f"  paperSize: {ps.paperSize}")
print(f"  fitToWidth: {ps.fitToWidth}, fitToHeight: {ps.fitToHeight}")
print(f"  print_area: {ws.print_area}")
print(f"  print_title_rows: {ws.print_title_rows}")
print(f"  margins: L={ws.page_margins.left} R={ws.page_margins.right}")

# ---- Hoja APU 01 ----
print()
print("=" * 70)
print("HOJA APU 01 - Estructura")
print("=" * 70)
ws = wb["APU 01"]
print(f"Dimensiones: {ws.dimensions}")
for r in range(1, min(ws.max_row, 35) + 1):
    vals = []
    for c in range(1, min(ws.max_column, 8) + 1):
        cell = ws.cell(row=r, column=c)
        if cell.value is not None:
            vals.append(f"{cell.coordinate}={repr(cell.value)}")
    if vals:
        print(f"  Fila {r}: " + " | ".join(vals))