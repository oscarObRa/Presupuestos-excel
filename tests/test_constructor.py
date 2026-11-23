"""Test end-to-end del constructor."""
import json, os, sys
sys.path.insert(0, ".")

import openpyxl
import pandas as pd

from src.indexador import Indexador
from src.lector_entrada import leer_entrada
from src.matcher import Matcher
from src.constructor import construir_presupuesto

os.makedirs("tests", exist_ok=True)

# 1. Catálogo mock realista
mock = {"conceptos": []}
conceptos_reales = [
    ("CAT-001", "EDIFICACION", "Preliminares", "Trazo", "Trazo y nivelacion de terreno para desplante de estructuras", "M2", 85.50),
    ("CAT-002", "EDIFICACION", "Preliminares", "Limpieza", "Limpieza y deshierbe de terreno con herramienta manual", "M2", 45.00),
    ("CAT-003", "EDIFICACION", "Cimentacion", "Excavacion", "Excavacion en cepa de 0 a 2m de profundidad material tipo B", "M3", 180.00),
    ("CAT-004", "EDIFICACION", "Cimentacion", "Plantilla", "Plantilla de concreto f'c=100 kg/cm2 de 5cm espesor", "M2", 120.00),
    ("CAT-005", "EDIFICACION", "Cimentacion", "Acero", "Acero de refuerzo en cimentacion fy=4200 kg/cm2 incluye habilitado y armado", "KG", 55.00),
    ("CAT-006", "EDIFICACION", "Estructura", "Concreto", "Concreto en columnas f'c=250 kg/cm2 hecho en obra", "M3", 3200.00),
    ("CAT-007", "EDIFICACION", "Estructura", "Cimbra", "Cimbra acabado aparente en columnas con madera de pino", "M2", 250.00),
    ("CAT-008", "EDIFICACION", "Estructura", "Acero", "Acero de refuerzo en columnas fy=4200 kg/cm2", "KG", 58.00),
    ("CAT-009", "EDIFICACION", "Albanileria", "Muro", "Muro de block hueco de concreto 15x20x40cm asentado con mortero", "M2", 350.00),
    ("CAT-010", "EDIFICACION", "Albanileria", "Aplanado", "Aplanado fino en muros con mortero 1:4 acabado pulido", "M2", 95.00),
]
for i, (clave, paq, n1, n2, desc, und, pu) in enumerate(conceptos_reales, 1):
    mock["conceptos"].append({
        "id": str(i), "paquete": paq, "nivel_1": n1, "nivel_2": n2,
        "clave": clave, "descripcion": desc, "unidad": und,
        "precio_unitario": pu, "aplicacion": "general",
    })

with open("tests/cat_realista.json", "w", encoding="utf-8") as f:
    json.dump(mock, f)

# 2. Entrada que coincide
entrada_data = [
    ("CONC-01", "PRELIMINARES", "Trazo y nivelacion de terreno para desplante de estructuras", "M2", 250.0),
    ("CONC-02", "PRELIMINARES", "Limpieza y deshierbe de terreno con herramienta manual", "M2", 400.0),
    ("CONC-03", "CIMENTACION", "Excavacion en cepa de 0 a 2m de profundidad material tipo B", "M3", 80.0),
    ("CONC-04", "CIMENTACION", "Plantilla de concreto f'c=100 kg/cm2 de 5cm espesor", "M2", 150.0),
    ("CONC-05", "CIMENTACION", "Acero de refuerzo en cimentacion fy=4200 kg/cm2 incluye habilitado y armado", "KG", 1200.0),
    ("CONC-06", "ESTRUCTURA", "Concreto en columnas f'c=250 kg/cm2 hecho en obra", "M3", 25.0),
    ("CONC-07", "ESTRUCTURA", "Cimbra acabado aparente en columnas con madera de pino", "M2", 180.0),
    ("CONC-08", "ESTRUCTURA", "Acero de refuerzo en columnas fy=4200 kg/cm2", "KG", 800.0),
    ("CONC-09", "ALBANILERIA", "Muro de block hueco de concreto 15x20x40cm asentado con mortero", "M2", 500.0),
    ("CONC-10", "ALBANILERIA", "Aplanado fino en muros con mortero 1:4 acabado pulido", "M2", 600.0),
]
df = pd.DataFrame(entrada_data, columns=["id", "partida", "concepto", "unidad", "cantidad"])
df.to_excel("tests/entrada_realista.xlsx", index=False)

# 3. Indexar
idx = Indexador("tests/cat_realista.json")
total = idx.cargar()
print(f"Catálogo: {total} conceptos")

# 4. Leer entrada
conceptos = leer_entrada("tests/entrada_realista.xlsx")
print(f"Entrada: {len(conceptos)} conceptos")

# 5. Matchear
m = Matcher(idx, umbral=75)
resultados = m.emparejar(conceptos)
stats = m.stats_matcheo(resultados)
print(f"Matcheo: {stats['exactos']} exactos, {stats['fuzzy']} fuzzy, {stats['sin_match']} sin match")
print(f"Confirmados: {stats['confirmados']} ({stats['pct_confirmados']}%)")

# Mostrar resultados
for r in resultados:
    print(f"  {r['id_entrada']:8s} | {r['estado']:11s} | {r['metodo']:7s} | "
          f"score={r['score']:5.1f} | {r['clave_sugerida']} | PU=${r['pu_sugerido']}")

# 6. Construir presupuesto
print("\n=== CONSTRUYENDO PRESUPUESTO ===\n")

proyecto = {
    "nombre": "AEROPUERTO INTERNACIONAL - SALA VIP",
    "clave": "PB-OMA-R01",
    "cliente": "OMA / AEROPUERTO INTERNACIONAL, S.A. DE C.V.",
    "fecha": "2026-11-20",
    "indirectos_campo": 0.14,
    "indirectos_oficina": 0.05,
    "financiamiento": 0.0135,
    "utilidad": 0.08,
    "iva": 0.16,
}

ruta_salida = construir_presupuesto(
    ruta_plantilla="plantillas/machote.xlsx",
    proyecto=proyecto,
    resultados_matcheo=resultados,
    indice=idx,
    carpeta_salida="salidas",
)

print(f"\n{'='*60}")
print(f"PRESUPUESTO GENERADO: {ruta_salida}")
print(f"Hojas generadas: {[s for s in openpyxl.load_workbook(ruta_salida).sheetnames if s not in ['Parametros', 'Categorias', 'Calculo_categoria', 'Cuadrillas', 'Notas']]}")
print(f"{'='*60}")