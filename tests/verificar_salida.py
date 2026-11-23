"""Verifica el presupuesto generado."""
import glob, os
import openpyxl

carpetas = sorted(glob.glob("salidas/*"))
ruta = os.path.join(carpetas[-1], "presupuesto.xlsx")
wb = openpyxl.load_workbook(ruta)

print("HOJAS:", wb.sheetnames)

# Presupuesto
ws = wb["Presupuesto"]
print("\n=== PRESUPUESTO (fila 7-19) ===")
for r in range(7, 20):
    a = ws.cell(row=r, column=1).value
    b = ws.cell(row=r, column=2).value
    e = ws.cell(row=r, column=5).value
    f = ws.cell(row=r, column=6).value
    g = ws.cell(row=r, column=7).value
    if a or b:
        print(f"  Fila {r}: partida={a} | clave={b} | cant={e} | PU={f} | importe={g}")

# Insumos
ws2 = wb["Insumos"]
print("\n=== INSUMOS (primeros 3) ===")
for r in range(5, 8):
    a = ws2.cell(row=r, column=1).value
    b = ws2.cell(row=r, column=2).value
    e = ws2.cell(row=r, column=5).value
    print(f"  {a} | {b} | costo={e}")

# Variables
ws3 = wb["Variables"]
print("\n=== VARIABLES ===")
for r in range(5, 22):
    a = ws3.cell(row=r, column=1).value
    b = ws3.cell(row=r, column=2).value
    if a:
        print(f"  {str(a)[:50]} -> {b}")

# Explosion
ws4 = wb["Explosion"]
print("\n=== EXPLOSION (última fila) ===")
mr = ws4.max_row
print(f"  Fila {mr}: {ws4.cell(row=mr, column=1).value} | total={ws4.cell(row=mr, column=7).value}")