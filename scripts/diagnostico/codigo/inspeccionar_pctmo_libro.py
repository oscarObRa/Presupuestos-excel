# -*- coding: utf-8 -*-
"""Inspección de las líneas %MO en el libro de verificación de VENADO."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(RAIZ))
import openpyxl

wb = openpyxl.load_workbook(RAIZ / "salidas" / "venado_revisado" / "presupuesto.xlsx", data_only=False)

ws = wb["CIM-01"]
print("=== CIM-01 (hoja APU) filas 1-20, cols A-G ===")
for r in range(1, 21):
    vals = [ws.cell(row=r, column=c).value for c in range(1, 8)]
    if any(v not in (None, "") for v in vals):
        print(f"r{r:>2} | " + " | ".join(str(v)[:34] if v is not None else "-" for v in vals))

print()
print("=== Insumos: filas INS-070 e INS-246 ===")
ins = wb["Insumos"]
for r in range(5, ins.max_row + 1):
    idl = ins.cell(row=r, column=1).value
    if idl in ("INS-070", "INS-246"):
        print(f"r{r} | " + " | ".join(str(ins.cell(row=r, column=c).value)[:42] for c in range(1, 10)))

print()
print("=== Todas las filas de hojas APU con '%' en unidad (muestra 12) ===")
n = 0
for nombre in wb.sheetnames:
    if nombre.startswith(("Presupuesto", "Portada", "Insumos", "Cuadrillas", "Variables", "Parametros", "Calculo", "Explosion", "Revision")):
        continue
    s = wb[nombre]
    for r in range(1, s.max_row + 1):
        uni = str(s.cell(row=r, column=3).value or "")
        if "%" in uni:
            n += 1
            if n <= 12:
                print(f"{nombre}!r{r} | A={s.cell(row=r, column=1).value} | B={str(s.cell(row=r, column=2).value)[:28]} "
                      f"| C={uni} | D={s.cell(row=r, column=4).value} | E={s.cell(row=r, column=5).value}")
print(f"total filas %: {n}")
wb.close()
