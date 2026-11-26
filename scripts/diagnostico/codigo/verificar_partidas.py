"""Verifica la estructura de partidas y subtotales en el presupuesto generado."""
import openpyxl

ruta = "scripts/diagnostico/salidas/vent_venado/presupuesto.xlsx"
wb = openpyxl.load_workbook(ruta)
ws = wb["Presupuesto"]

print("=== ESTRUCTURA DE LA HOJA PRESUPUESTO ===\n")
for r in range(7, ws.max_row + 1):
    a = ws.cell(row=r, column=1).value
    b = ws.cell(row=r, column=2).value
    e = ws.cell(row=r, column=5).value
    g = ws.cell(row=r, column=7).value

    # Detectar tipo de fila
    if a and str(a).startswith("PARTIDA"):
        print(f"\n[PARTIDA] Fila {r}: {str(a)[:70]}")
    elif a and str(a).startswith("Subtotal"):
        print(f"  >> SUBTOTAL (fila {r}): {g}")
    elif a and ("SUBTOTAL GENERAL" in str(a) or "Indirectos" in str(a) or "Utilidad" in str(a)
                or "TOTAL" in str(a) or "IVA" in str(a) or "GRAN" in str(a)):
        print(f"\n*** {a}  =>  {g}")
    elif b:  # fila de concepto
        print(f"    {b:8s} | cant={e} | importe={g}")
    elif a:
        print(f"\n{a}  =>  {g}")

print(f"\nTotal filas: {ws.max_row}")