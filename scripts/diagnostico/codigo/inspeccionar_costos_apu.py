# -*- coding: utf-8 -*-
"""Costos reales de los insumos del APU mal matcheado + verificación del PU."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
import openpyxl
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[3]
RUTA = RAIZ / "salidas" / "2026-11-27_1451" / "presupuesto.xlsx"

wb = openpyxl.load_workbook(RUTA, data_only=False)
ins = wb["Insumos"]
print("=== Insumos del APU mal matcheado (M-01..M-07 = K78.7) ===")
objetivo = {"INS-117", "INS-199", "INS-152", "INS-077", "INS-036", "INS-033", "INS-021"}
for r in range(5, ins.max_row + 1):
    idl = ins.cell(row=r, column=1).value
    if idl in objetivo:
        # A=Id B=Tipo C=Desc D=Unidad E=Costo F=Cant G=Importe
        print(f"{idl} | {ins.cell(row=r, column=2).value} | "
              f"{str(ins.cell(row=r, column=3).value)[:45]} | "
              f"{ins.cell(row=r, column=4).value} | costo={ins.cell(row=r, column=5).value} | "
              f"cant={ins.cell(row=r, column=6).value}")

# Costos de cuadrilla en hoja Cuadrillas
cu = wb["Cuadrillas"]
print()
print("=== Cuadrillas (A..I) filas con INS- ===")
for r in range(4, min(cu.max_row, 40) + 1):
    idl = cu.cell(row=r, column=1).value
    if idl and str(idl).startswith("INS-"):
        print(f"{idl} | {str(cu.cell(row=r, column=2).value)[:40]} | $/Jornal={cu.cell(row=r, column=9).value}")

# PU OFICIAL de K78.7 impreso en la hoja APU
m01 = wb["M-01"]
print()
print("=== M-01 resumen (r15-r19) ===")
for r in range(15, 20):
    print(f"r{r}:", m01.cell(row=r, column=3).value, "|", m01.cell(row=r, column=6).value)
wb.close()

# Reproducir el PU con la misma aritmetica del APU
# (necesita costos resueltos; si el VLOOKUP no resuelve, el IFERROR usa el costo de la matriz)
import sqlite3
DB = r"C:\Pontem_dev\costeo\dist\contrato_costos\1.2.0\costos_canonico.db"
conn = sqlite3.connect("file:" + DB.replace("\\", "/") + "?mode=ro", uri=True)
conn.row_factory = sqlite3.Row
print()
print("=== Esquema apus/apu_lineas ===")
for t in ("apus", "apu_lineas"):
    cols = [c[1] for c in conn.execute(f"PRAGMA table_info({t})")]
    print(t, "->", cols)
conn.close()
