"""Test del lector con partidas sobre VENADO.xlsx"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[3]))  # raiz del repo
import sys

from src.lector_entrada import leer_entrada, validar_entrada

ruta = r"H:\Mi unidad\ICEL\PUs\CASA VENADO\VENADO.xlsx"

v = validar_entrada(ruta)
print("=== VALIDACION ===")
print(f"  Valido: {v['valido']}")
print(f"  Conceptos: {v['num_filas']}")
print(f"  Partidas: {v['num_partidas']}")
print(f"  Error: {v['error']}")

conceptos = leer_entrada(ruta)
print(f"\n=== PRIMEROS 5 CONCEPTOS ===")
for c in conceptos[:5]:
    print(f"  [{c['partida'][:35]:35s}] {c['id']:8s} | {c['concepto'][:40]} | {c['unidad']} | {c['cantidad']}")

print(f"\n=== CONCEPTOS POR PARTIDA ===")
from collections import Counter
counts = Counter(c["partida"] for c in conceptos)
for p, n in counts.items():
    print(f"  {p[:60]:60s} : {n} conceptos")

print(f"\nTotal conceptos: {len(conceptos)}")