"""Test de subpartidas: 2 niveles de agrupación."""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[3]))  # raiz del repo
import json, sys
from pathlib import Path

import pandas as pd
from src.lector_entrada import leer_entrada
from src.indexador import Indexador
from src.matcher import Matcher
from src.constructor import construir_presupuesto

# Archivo con PARTIDA + SUBPARTIDA incrustadas
filas = [
    ("PARTIDA 01: CIMENTACION", None, None, None),
    ("CIM-01", "Excavacion en cepa", "m3", 10.0),
    ("SUBPARTIDA 01.1: ZAPATAS", None, None, None),
    ("CIM-02", "Zapata aislada ZA-1", "pza", 4.0),
    ("CIM-03", "Zapata aislada ZA-2", "pza", 6.0),
    ("SUBPARTIDA 01.2: LOSA", None, None, None),
    ("CIM-04", "Losa de cimentacion LC-15", "m2", 40.0),
    ("PARTIDA 02: MUROS", None, None, None),
    ("MUR-01", "Muro de block 15cm", "m2", 80.0),
    ("MUR-02", "Muro de block 20cm", "m2", 20.0),
]
df = pd.DataFrame(filas, columns=["ID", "Concepto", "Unidad", "Cantidad"])
df.to_excel("scripts/diagnostico/datos/entrada_subpartidas.xlsx", index=False)

conceptos = leer_entrada("scripts/diagnostico/datos/entrada_subpartidas.xlsx")
print("=== CONCEPTOS LEIDOS ===")
for c in conceptos:
    print(f"  partida={c['partida'][:25]:25s} | sub={str(c['subpartida'])[:25]:25s} | {c['id']}")

# Catalogo mock
mock = {"conceptos": []}
for i, c in enumerate(conceptos, 1):
    mock["conceptos"].append({
        "id": str(i), "paquete": "X", "nivel_1": "Y", "nivel_2": "Z",
        "clave": c["id"], "descripcion": c["concepto"],
        "unidad": c["unidad"], "precio_unitario": 100.0 * i, "aplicacion": "general",
    })
Path("scripts/diagnostico/datos/catalogo_sub.json").write_text(json.dumps(mock), encoding="utf-8")

idx = Indexador("scripts/diagnostico/datos/catalogo_sub.json")
idx.cargar()
m = Matcher(idx, umbral=85)
resultados = m.emparejar(conceptos)

proyecto = {"nombre": "TEST SUBPARTIDAS", "clave": "T", "cliente": "X",
            "fecha": "2026-11-21", "indirectos_campo": 0.14, "indirectos_oficina": 0.05,
            "financiamiento": 0.0135, "utilidad": 0.08, "iva": 0.16}

salida = construir_presupuesto(
    ruta_plantilla="plantillas/machote.xlsx",
    proyecto=proyecto, resultados_matcheo=resultados, indice=idx,
    carpeta_salida="scripts/diagnostico/salidas/vent_sub", usar_timestamp=False,
)
print(f"\nSalida: {salida}")

# Verificar estructura
import openpyxl
ws = openpyxl.load_workbook(str(salida))["Presupuesto"]
print("\n=== ESTRUCTURA PRESUPUESTO ===")
for r in range(8, ws.max_row + 1):
    a = ws.cell(row=r, column=1).value
    b = ws.cell(row=r, column=2).value
    g = ws.cell(row=r, column=7).value
    if a:
        print(f"  Fila {r}: A='{a}'  G={g}")
    elif b:
        print(f"  Fila {r}: concepto {b}")