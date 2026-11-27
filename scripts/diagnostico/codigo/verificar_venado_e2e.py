# -*- coding: utf-8 -*-
"""Rebuild end-to-end de CASA VENADO con el matcher corregido y revisión del resultado."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(RAIZ))

from src.lector_entrada import leer_entrada
from src.costeo_client import ClienteCosteo
from src.matcher import Matcher
from src.apu_db import ProveedorAPU
from src.constructor import construir_presupuesto
import openpyxl
from collections import Counter

ENTRADA = r"H:\Mi unidad\ICEL\PUs\CASA VENADO\VENADO.xlsx"
SALIDA = RAIZ / "salidas" / "venado_revisado"

conceptos = leer_entrada(ENTRADA)
print(f"Entrada: {len(conceptos)} conceptos")

cliente = ClienteCosteo(r"C:\Pontem_dev\costeo")
matcher = Matcher(cliente, umbral=55)
resultados = matcher.emparejar(conceptos)

estados = Counter(r["estado"] for r in resultados)
print("Estados:", dict(estados))

# El caso del usuario: M-0x no debe volver a caer en K78.7
m01 = next(r for r in resultados if r["id_entrada"] == "M-01")
print(f"M-01 -> estado={m01['estado']} clave={m01.get('clave_sugerida') or '-'} "
      f"score={m01.get('score')} desc={str(m01.get('descripcion_sugerida'))[:60]}")

prov = ProveedorAPU(repo=r"C:\Pontem_dev\costeo")
ruta = construir_presupuesto(
    ruta_plantilla=str(RAIZ / "plantillas" / "machote.xlsx"),
    proyecto={"nombre": "CASA VENADO (verificacion)", "clave": "VENADO-R02",
              "cliente": "ICEL", "fecha": "2026-11-27",
              "cotizacion": "COT-VENADO", "revision": "R02"},
    resultados_matcheo=resultados,
    carpeta_salida=str(SALIDA),
    usar_timestamp=False,
    apu_provider=prov,
    validar=False, recalcular=False,
)

# Inspección del libro resultante
wb = openpyxl.load_workbook(ruta, data_only=False)
ws = wb["Presupuesto"]
pues_altos = []
for r in range(1, ws.max_row + 1):
    f = ws.cell(row=r, column=6).value
    idv = ws.cell(row=r, column=1).value
    if isinstance(f, (int, float)) and f > 200000 and idv:
        pues_altos.append((r, idv, f))
print()
print(f"PUs estáticos > $200,000: {len(pues_altos)}")
for r, idv, f in pues_altos[:8]:
    print(f"  r{r} {idv} PU={f:,.2f} | {str(ws.cell(row=r, column=3).value)[:55]}")

# K78.7 no debe aparecer como clave de ningún concepto
con_k78 = [(r, ws.cell(row=r, column=1).value) for r in range(1, ws.max_row + 1)
           if str(ws.cell(row=r, column=2).value or "").strip() == "K78.7"]
print(f"Conceptos con clave K78.7: {len(con_k78)}")

# %MO no debe estar en Insumos
ins = wb["Insumos"]
pct_insumos = [r for r in range(5, ins.max_row + 1)
               if "%" in str(ins.cell(row=r, column=4).value or "")]
print(f"Líneas '%' en Insumos: {len(pct_insumos)}")

# Portada: anchos de columna restaurados
pt = wb["Portada"]
anchos = {c: pt.column_dimensions[c].width for c in "ABCDEF"}
print(f"Portada anchos A-F: {anchos}")
wb.close()
print()
print("Libro de verificación:", ruta)
