"""Verifica encabezado y WBS del nuevo Presupuesto."""
import openpyxl

ruta = "scripts/diagnostico/salidas/vent_venado/presupuesto.xlsx"
wb = openpyxl.load_workbook(ruta)
ws = wb["Presupuesto"]

print("=== ENCABEZADO (filas 1-6) ===")
for r in range(1, 7):
    vals = []
    for c in range(1, 9):
        v = ws.cell(row=r, column=c).value
        if v is not None:
            vals.append(f"{ws.cell(row=r,column=c).coordinate}={repr(v)[:60]}")
    print(f"  F{r}: " + " | ".join(vals) if vals else f"  F{r}: (vacía)")

print("\n=== ESTRUCTURA WBS (primeras 25 filas de datos) ===")
for r in range(7, 32):
    a = ws.cell(row=r, column=1).value
    b = ws.cell(row=r, column=2).value
    c3 = ws.cell(row=r, column=3).value
    d = ws.cell(row=r, column=4).value
    g = ws.cell(row=r, column=7).value
    h = ws.cell(row=r, column=8).value
    if a is None and b is None:
        continue
    print(f"  F{r}: A={a!s:8} | B={str(b)[:8]:8} | C={str(c3)[:35]:35} | D={d} | G={g} | H={h}")

print("\n=== PIE (últimas 12 filas) ===")
for r in range(ws.max_row - 11, ws.max_row + 1):
    vals = []
    for c in range(1, 9):
        v = ws.cell(row=r, column=c).value
        if v is not None:
            vals.append(f"{ws.cell(row=r,column=c).coordinate}={repr(v)[:55]}")
    if vals:
        print(f"  F{r}: " + " | ".join(vals))

print(f"\nPage setup: orientation={ws.page_setup.orientation}, fitToWidth={ws.page_setup.fitToWidth}, print_titles={ws.print_title_rows}")