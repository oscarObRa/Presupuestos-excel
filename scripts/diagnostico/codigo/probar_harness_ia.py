# -*- coding: utf-8 -*-
"""Ciclo completo del harness de IA con candidatos reales del matcher.

Matcher sobre una muestra real de CASA VENADO -> resolver() de la IA ->
aplicar_sugerencias() con el cliente real. Reporta sin exponer la llave.
"""
import sys
sys.stdout.reconfigure(encoding="utf-8")
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(RAIZ))

from src.lector_entrada import leer_entrada
from src.costeo_client import ClienteCosteo
from src.matcher import Matcher
from src.llm import cargar_config_llm, ClienteLLM, aplicar_sugerencias

ENTRADA = r"H:\Mi unidad\ICEL\PUs\CASA VENADO\VENADO.xlsx"
IDS = {"M-01", "M-02", "S-45", "CAR-03"}

conceptos = [c for c in leer_entrada(ENTRADA) if c["id"] in IDS]
cliente = ClienteCosteo(r"C:\Pontem_dev\costeo")
matcher = Matcher(cliente, umbral=55)
resultados = matcher.emparejar(conceptos)

print("=== Estados del matcher (muestra) ===")
for r in resultados:
    print(f"{r['id_entrada']}: {r['estado']} | score {r.get('score')} | "
          f"candidatos: {len(r.get('candidatos') or [])} | clave {r.get('clave_sugerida') or '-'}")

sin_candidatos = [r["id_entrada"] for r in resultados
                  if r["estado"] in ("dudoso", "sin_match") and not (r.get("candidatos") or [])]
print()
print(f"Conceptos dudosos/sin_match SIN candidatos (la IA nunca los ve): {sin_candidatos}")

cfg = cargar_config_llm("config.yaml")
cfg["max_tokens"] = 900
llm = ClienteLLM(cfg)
print()
print("IA:", llm.estado())

try:
    sugerencias = llm.resolver(resultados, buscar=matcher.candidatos_para_ia)
except RuntimeError as e:
    print("RUNTIME ( surfaced por el fix):", e)
    raise SystemExit(0)

print(f"Sugerencias de la IA: {len(sugerencias)}")
for s in sugerencias:
    print(f"  {s.get('id')}: clave={s.get('clave')} confianza={s.get('confianza')} "
          f"motivo={str(s.get('motivo'))[:60]}")

aplicadas = aplicar_sugerencias(resultados, cliente, sugerencias,
                                auto_aplicar=cfg.get("auto_aplicar", True),
                                confianza_minima=cfg.get("confianza_minima", 0.85))
print(f"Aplicadas al presupuesto: {aplicadas}")
for r in resultados:
    if r.get("clave_ia"):
        print(f"  {r['id_entrada']} -> {r['estado']} clave_ia={r['clave_ia']} "
              f"PU={r.get('pu_sugerido')}")
