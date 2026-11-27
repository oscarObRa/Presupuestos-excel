# -*- coding: utf-8 -*-
"""Diagnóstico del match malo M-01 -> K78.7 y del APU generado."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
import openpyxl
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(RAIZ))

RUTA = RAIZ / "salidas" / "2026-11-27_1451" / "presupuesto.xlsx"

# 1) Hoja APU 'M-01' en el libro generado
wb = openpyxl.load_workbook(RUTA, data_only=False)
if "M-01" in wb.sheetnames:
    ws = wb["M-01"]
    print("=== HOJA APU 'M-01' (del libro generado) ===")
    for r in range(1, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, 8)]
        if any(v is not None and str(v).strip() != "" for v in vals):
            print(f"r{r:>3} | " + " | ".join(str(v)[:38] if v is not None else "-" for v in vals))
wb.close()

# 2) Qué dice la BD canónica (SQLite) sobre el APU de K78.7
import sqlite3
DB = r"C:\Pontem_dev\costeo\dist\contrato_costos\1.2.0\costos_canonico.db"
conn = sqlite3.connect("file:" + DB.replace("\\", "/") + "?mode=ro", uri=True)
conn.row_factory = sqlite3.Row
tablas = [t[0] for t in conn.execute("SELECT name FROM sqlite_master WHERE type IN ('table','view')")]
print()
print("=== Tablas/vistas:", tablas)
for t in tablas:
    if "apu" in t.lower():
        filas = conn.execute(f"SELECT * FROM {t} WHERE clave_concepto='K78.7' LIMIT 20").fetchall()
        if filas:
            cols = filas[0].keys()
            print(f"--- {t} de K78.7 ({len(filas)} filas) cols={list(cols)}")
            for f in filas:
                print("   ", {k: f[k] for k in f.keys()})
conn.close()
