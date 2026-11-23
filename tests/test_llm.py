"""Test de llm.py sin red: parseo de JSON y aplicación de sugerencias."""
import sys
sys.path.insert(0, ".")

from src.llm import _limpiar_json
from src.costeo_client import ClienteCosteo
from src.indexador import Indexador
from src.llm import aplicar_sugerencias

# 1. Limpieza de JSON con cercas
assert _limpiar_json('```json\n{"resultados":[]}\n```') == '{"resultados":[]}'
print("OK limpiar_json")

cli = ClienteCosteo()
idx = Indexador(str(cli.indice_path)); idx.cargar()

# Tomar un concepto real del índice para simular sugerencia
concepto = idx.todos()[0]
clave = concepto["clave"]

# Resultado simulado "dudoso" con ese candidato en su lista
resultados = [{
    "id_entrada": "X-01", "concepto_entrada": "algo", "unidad_entrada": "m2",
    "cantidad": 10, "clave_sugerida": "", "descripcion_sugerida": "",
    "unidad_sugerida": "", "pu_sugerido": 0, "score": 20.0,
    "metodo": "sin_match", "estado": "dudoso",
    "candidatos": [{"clave": clave, "descripcion": concepto["descripcion"],
                    "unidad": concepto["unidad"],
                    "precio_unitario": concepto["precio_unitario"]}],
}]

sug = [{"id": "X-01", "clave": clave, "confianza": 0.9, "motivo": "coincide"}]
n = aplicar_sugerencias(resultados, idx, sug, auto_aplicar=True)
r = resultados[0]
print("aplicadas:", n)
print("estado:", r["estado"], "| clave:", r["clave_sugerida"], "| pu:", r["pu_sugerido"])
assert n == 1 and r["estado"] == "confirmado_ia" and r["clave_sugerida"] == clave

# Clave inventada NO debe aplicarse
resultados2 = [dict(resultados[0], estado="dudoso", clave_sugerida="", clave_ia=None)]
sug2 = [{"id": "X-01", "clave": "INVENTADA-XYZ", "confianza": 0.9, "motivo": "?"}]
n2 = aplicar_sugerencias(resultados2, idx, sug2, auto_aplicar=True)
print("aplicadas con clave inventada (debe ser 0):", n2)
assert n2 == 0

print("\nTODO OK — lógica de IA validada")