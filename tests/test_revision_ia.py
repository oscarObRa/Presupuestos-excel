"""Verifica columnas IA en la hoja Revisión."""
import sys
sys.path.insert(0, ".")
from src.constructor import Constructor
from src.indexador import Indexador

idx = Indexador("tests/catalogo_costos.json"); idx.cargar()
c1 = idx.todos()[0]
c2 = idx.todos()[1]

resultados = [
    {"id_entrada": "A-1", "partida": "P1", "subpartida": None,
     "concepto_entrada": "x1", "unidad_entrada": c1["unidad"], "cantidad": 1,
     "clave_sugerida": c1["clave"], "descripcion_sugerida": c1["descripcion"],
     "unidad_sugerida": c1["unidad"], "pu_sugerido": c1["precio_unitario"],
     "score": 99, "metodo": "exacto", "estado": "confirmado", "candidatos": []},
    {"id_entrada": "A-2", "partida": "P1", "subpartida": None,
     "concepto_entrada": "x2", "unidad_entrada": c2["unidad"], "cantidad": 2,
     "clave_sugerida": "", "descripcion_sugerida": "", "unidad_sugerida": "",
     "pu_sugerido": 0, "score": 20, "metodo": "sin_match", "estado": "sin_match",
     "candidatos": [{"clave": c2["clave"], "descripcion": c2["descripcion"],
                     "unidad": c2["unidad"], "precio_unitario": c2["precio_unitario"]}],
     "clave_ia": c2["clave"], "confianza_ia": 0.88, "motivo_ia": "coincide por unidad y tipo"},
]

proyecto = {"nombre": "TEST IA", "cliente": "X", "fecha": "2026-11-21",
            "cotizacion": "C-1", "revision": "R1",
            "indirectos_campo": 0.14, "indirectos_oficina": 0.05,
            "financiamiento": 0.0135, "utilidad": 0.08, "iva": 0.16}

c = Constructor("plantillas/machote.xlsx", "tests/vent_ia", usar_timestamp=False)
c.copiar_plantilla(); c.escribir_variables(proyecto)
c.escribir_insumos(resultados, idx); c.escribir_revision(resultados)
c.escribir_presupuesto(resultados, proyecto); c.escribir_explosion(resultados)
ruta = c.guardar()

import openpyxl
ws = openpyxl.load_workbook(str(ruta))["Revision"]
headers = [ws.cell(row=4, column=c).value for c in range(1, 14)]
print("Headers Revision:", headers)
assert headers[10] == "Clave IA" and headers[12] == "Motivo IA"
print("Fila A-2 Clave IA:", ws.cell(row=6, column=11).value)
print("Fila A-2 Motivo IA:", ws.cell(row=6, column=13).value)
print("OK — Revisión con columnas IA")