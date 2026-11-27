# -*- coding: utf-8 -*-
"""Reproduce el matcheo de los conceptos M-0x de VENADO y muestra los candidatos."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(RAIZ))

from src.lector_entrada import leer_entrada
from src.costeo_client import ClienteCosteo
from src.matcher import Matcher
from src.apu_db import ProveedorAPU
from src.indexador import Indexador

entrada = RAIZ.parent / "costeo"  # no aplica
entrada_path = r"H:\Mi unidad\ICEL\PUs\CASA VENADO\VENADO.xlsx"

conceptos = leer_entrada(entrada_path)
objetivos = [c for c in conceptos if str(c["id"]).upper() in
             ("M-01", "M-02", "M-03", "M-04", "M-05", "M-06", "M-07")]
print(f"Conceptos objetivo: {[c['id'] for c in objetivos]}")

cliente = ClienteCosteo(r"C:\Pontem_dev\costeo")
print("Cliente:", cliente.descripcion_contrato())

matcher = Matcher(cliente, umbral=55)

for c in objetivos:
    print()
    print(f"--- {c['id']} | {c['concepto'][:80]}...")
    res = matcher.emparejar([c])[0]
    print(f"   estado={res.get('estado')} metodo={res.get('metodo')} "
          f"clave={res.get('clave_sugerida')} score={res.get('score')}")
    print(f"   descripcion_sugerida: {str(res.get('descripcion_sugerida'))[:90]}")
    print(f"   pu_sugerido: {res.get('pu_sugerido')}")
    # top candidatos crudos del cliente
    cand = cliente.buscar(c["concepto"], unidad=c.get("unidad"), limite=5)
    for r in cand[:5]:
        print(f"     [score {r.get('score')}] {r.get('clave')} | "
              f"desc={str(r.get('descripcion'))[:40]} | nivel2={str(r.get('nivel_2'))[:40]}")
