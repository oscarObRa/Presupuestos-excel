# -*- coding: utf-8 -*-
"""Inspecciona el libro generado de CASA VENADO: matches raros y PUs sospechosos."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
import openpyxl

RUTA = r"salidas\2026-11-27_1451\presupuesto.xlsx"
wb = openpyxl.load_workbook(RUTA, data_only=False)

ws = wb["Presupuesto"]
print("=== PRESUPUESTO: filas con PU estatico > 100k ===")
for r in range(1, ws.max_row + 1):
    f = ws.cell(row=r, column=6).value
    if isinstance(f, (int, float)) and f > 100000:
        print(f"r{r} | {ws.cell(row=r, column=1).value} | "
              f"{str(ws.cell(row=r, column=3).value)[:70]} | PU: {f}")

print()
print("=== PRESUPUESTO: filas que mencionan diametro / 12 pulgadas ===")
for r in range(1, ws.max_row + 1):
    c = str(ws.cell(row=r, column=3).value or "")
    b = str(ws.cell(row=r, column=2).value or "")
    low = (b + " " + c).lower()
    if "diam" in low or '12"' in low or "12 pulg" in low or "12in" in low:
        print(f"r{r} | {ws.cell(row=r, column=1).value} | clave {b} | "
              f"{c[:70]} | PU: {ws.cell(row=r, column=6).value}")

print()
print("=== PRESUPUESTO: filas de PU formula (ligadas a APU) ===")
n = 0
for r in range(1, ws.max_row + 1):
    f = ws.cell(row=r, column=6).value
    if isinstance(f, str) and f.startswith("='"):
        n += 1
        if n <= 8:
            print(f"r{r} | {ws.cell(row=r, column=1).value} | {f}")
print(f"total formulas PU: {n}")

print()
print("=== REVISION: matches de conceptos con 'muebl' ===")
if "Revision" in wb.sheetnames:
    rv = wb["Revision"]
    for r in range(5, rv.max_row + 1):
        idv = rv.cell(row=r, column=1).value
        con = str(rv.cell(row=r, column=2).value or "")
        if "muebl" in con.lower() or "muebl" in str(idv or "").lower():
            sug = rv.cell(row=r, column=6).value
            score = rv.cell(row=r, column=7).value
            estado = rv.cell(row=r, column=9).value
            pu_of = rv.cell(row=r, column=11).value
            pu_apu = rv.cell(row=r, column=12).value
            print(f"r{r} | {idv} | entrada: {con[:55]} | sugerida: {str(sug)[:45]} "
                  f"| score {score} | {estado} | PUof {pu_of} | PUapu {pu_apu}")
wb.close()
