"""Inspección de colores (theme), imágenes y estructura APU del mockup."""
import openpyxl
from openpyxl.utils import get_column_letter

ruta = r"H:\Mi unidad\ICEL\PUs\Pretil OMA MTY\Generador_APU_mockup.xlsx"
wb = openpyxl.load_workbook(ruta, data_only=False)

def color_info(color):
    if color is None:
        return "none"
    if color.type == "rgb":
        return f"rgb={color.rgb}"
    if color.type == "theme":
        return f"theme={color.theme},tint={color.tint}"
    if color.type == "indexed":
        return f"indexed={color.indexed}"
    return str(color.type)

def fill_info(cell):
    f = cell.fill
    if f is None or f.patternType is None:
        return None
    return f"pattern={f.patternType} fg[{color_info(f.fgColor)}] bg[{color_info(f.bgColor)}]"

def font_info(cell):
    fo = cell.font
    return f"name={fo.name} sz={fo.size} bold={fo.bold} color[{color_info(fo.color)}]"

print("=" * 60)
print("PResupuesto — colores por tipo de fila")
print("=" * 60)
ws = wb["PResupuesto"]
for r in [5, 6, 7, 8, 9, 10, 11, 22, 23, 47, 50, 52]:
    for c in [1, 2, 3, 7, 10]:
        cell = ws.cell(row=r, column=c)
        fi = fill_info(cell)
        if fi or cell.value is not None:
            print(f"  {cell.coordinate}: {font_info(cell)} | {fi}")

print()
print("=== Imágenes en PROPUESTA ===")
wp = wb["PROPUESTA"]
print(f"  Imágenes: {len(wp._images)}")
for img in wp._images:
    try:
        print(f"    anchor={img.anchor._from.col},{img.anchor._from.row} size={img.width}x{img.height}")
    except Exception as e:
        print(f"    (img {e})")

print()
print("=== Imágenes en PResupuesto ===")
print(f"  Imágenes: {len(ws._images)}")

print()
print("=== APU 01 estructura completa ===")
wa = wb["APU 01"]
print(f"  Dimensiones: {wa.dimensions}, Max fila: {wa.max_row}")
print(f"  Anchos: {[(k, v.width) for k, v in wa.column_dimensions.items() if v.width]}")
print(f"  Merged: {list(wa.merged_cells.ranges)[:15]}")
for r in range(1, wa.max_row + 1):
    vals = []
    for c in range(1, min(wa.max_column, 8) + 1):
        cell = wa.cell(row=r, column=c)
        if cell.value is not None:
            vals.append(f"{cell.coordinate}={repr(cell.value)[:40]}")
    if vals:
        print(f"  F{r}: " + " | ".join(vals))