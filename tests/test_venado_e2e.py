"""Test end-to-end del flujo con partidas usando VENADO.xlsx"""
import json, sys
from pathlib import Path
sys.path.insert(0, ".")

from src.lector_entrada import leer_entrada
from src.indexador import Indexador
from src.matcher import Matcher
from src.constructor import construir_presupuesto

ruta_venado = r"H:\Mi unidad\ICEL\PUs\CASA VENADO\VENADO.xlsx"
conceptos = leer_entrada(ruta_venado)
print(f"Conceptos leidos: {len(conceptos)}")

# Construir catalogo mock a partir de los conceptos reales (para que matchee exacto)
mock = {"conceptos": []}
for i, c in enumerate(conceptos, 1):
    mock["conceptos"].append({
        "id": str(i), "paquete": c["partida"].split(":")[0],
        "nivel_1": "CONSTRUCCION", "nivel_2": c["id"],
        "clave": c["id"],
        "descripcion": c["concepto"],
        "unidad": c["unidad"],
        "precio_unitario": round(100 + i * 37.5, 2),
        "aplicacion": "general",
    })

cat_path = Path("tests/catalogo_venado.json")
cat_path.write_text(json.dumps(mock, ensure_ascii=False), encoding="utf-8")
print(f"Catalogo mock: {len(mock['conceptos'])} conceptos")

idx = Indexador(str(cat_path))
idx.cargar()

m = Matcher(idx, umbral=85)
resultados = m.emparejar(conceptos)
stats = m.stats_matcheo(resultados)
print(f"Matcheo: {stats['exactos']} exactos, {stats['fuzzy']} fuzzy, {stats['sin_match']} sin match")

proyecto = {
    "nombre": "CASA VENADO",
    "clave": "VENADO",
    "cliente": "CLIENTE ICEL",
    "fecha": "2026-11-21",
    "indirectos_campo": 0.14,
    "indirectos_oficina": 0.05,
    "financiamiento": 0.0135,
    "utilidad": 0.08,
    "iva": 0.16,
}

salida = construir_presupuesto(
    ruta_plantilla="plantillas/machote.xlsx",
    proyecto=proyecto,
    resultados_matcheo=resultados,
    indice=idx,
    carpeta_salida="tests/vent_venado",
    usar_timestamp=False,
)
print(f"\nSalida: {salida}")