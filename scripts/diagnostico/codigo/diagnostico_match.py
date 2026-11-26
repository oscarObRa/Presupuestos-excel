"""Diagnóstico: comparar buscador oficial vs rapidfuzz sobre VENADO."""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[3]))  # raiz del repo
import sys, json
sys.path.insert(0, r"C:\Pontem_dev\costeo")

from src.lector_entrada import leer_entrada
from src.matcher import normalizar

# 1. Buscador oficial
import buscador_costos as bc

ruta = r"H:\Mi unidad\ICEL\PUs\CASA VENADO\VENADO.xlsx"
conceptos = leer_entrada(ruta)

# tomar 6 ejemplos representativos
muestras = [conceptos[i] for i in [0, 2, 4, 6, 11, 12]]

from rapidfuzz import fuzz, process

# precargar descripciones para rapidfuzz
meta, cs = bc.cargar(ruta=r"C:\Pontem_dev\costeo\costos_unitarios\catalogo_costos.json")
descs = [c["_desc_norm"] for c in cs]

for m in muestras:
    texto = m["concepto"]
    print("=" * 90)
    print("ENTRADA:", texto[:80])
    print("  UNIDAD:", m["unidad"])

    # Oficial
    res = bc.buscar(texto, limite=3, ruta_indice=r"C:\Pontem_dev\costeo\costos_unitarios\catalogo_costos.json")
    print("  [Oficial]")
    for r in res:
        print(f"    score={r['score']:.3f} | {r['clave']} | {r['unidad']} | ${r['precio_unitario']} | {r['descripcion'][:60]}")

    # rapidfuzz scorers
    q = normalizar(texto)
    for scorer_name, scorer in [("token_set_ratio", fuzz.token_set_ratio),
                                 ("partial_ratio", fuzz.partial_ratio),
                                 ("WRatio", fuzz.WRatio)]:
        best = process.extractOne(q, descs, scorer=scorer)
        if best:
            desc_match, score, idx = best
            c = cs[idx]
            print(f"  [rapidfuzz {scorer_name:16}] score={score:5.1f} | {c['clave']} | {c['unidad']} | {c['descripcion'][:55]}")
    print()
