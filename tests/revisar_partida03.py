"""Revisa partida 03 (ENT-*) para detectar el importe atipico."""
import sys, os
sys.path.insert(0, ".")
from src.indexador import Indexador
from src.lector_entrada import leer_entrada
from src.matcher import Matcher

idx = Indexador(r"C:\Pontem_dev\costeo\costos_unitarios\catalogo_costos.json")
idx.cargar()
conceptos = leer_entrada(r"H:\Mi unidad\ICEL\PUs\CASA VENADO\VENADO.xlsx")
m = Matcher(idx, umbral=55, piso=30)
res = m.emparejar(conceptos)

for r in res:
    if r["clave_sugerida"] and r["clave_sugerida"].startswith(""):
        if r["id_entrada"].startswith("ENT-"):
            imp = r["cantidad"] * (r["pu_sugerido"] or 0)
            print(f"{r['id_entrada']:8s} cant={r['cantidad']:>8.2f} PU=${r['pu_sugerido']:>12,.2f} imp=${imp:>14,.2f} | {r['clave_sugerida']:12s} | {r['descripcion_sugerida'][:45]} | score={r['score']}")
