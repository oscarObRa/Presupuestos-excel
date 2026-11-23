"""Test de matcher con coincidencias exactas."""
import json, os, sys
sys.path.insert(0, ".")

from src.indexador import Indexador
from src.matcher import Matcher

os.makedirs("tests", exist_ok=True)

# Catálogo con conceptos realistas
mock = {
    "conceptos": [
        {"id": "1", "paquete": "A", "nivel_1": "B", "nivel_2": "C",
         "clave": "CAT-099", "descripcion": "Trazo y nivelacion de terreno",
         "unidad": "M2", "precio_unitario": 500.0, "aplicacion": "general"},
        {"id": "2", "paquete": "A", "nivel_1": "B", "nivel_2": "C",
         "clave": "CAT-100", "descripcion": "Muro de block hueco de concreto 15x20x40cm",
         "unidad": "M2", "precio_unitario": 350.0, "aplicacion": "general"},
        {"id": "3", "paquete": "A", "nivel_1": "B", "nivel_2": "C",
         "clave": "CAT-101", "descripcion": "Acero de refuerzo fy=4200 kg/cm2",
         "unidad": "KG", "precio_unitario": 45.0, "aplicacion": "general"},
    ]
}
json.dump(mock, open("tests/cat_realista.json", "w", encoding="utf-8"))

idx = Indexador("tests/cat_realista.json")
idx.cargar()
m = Matcher(idx, umbral=85)

# Entrada: 2 exactos + 1 inexistente
entrada = [
    {"id": "C01", "partida": "TEST", "concepto": "Trazo y nivelación de terreno", "unidad": "M2", "cantidad": 100},
    {"id": "C02", "partida": "TEST", "concepto": "Muro de block hueco de concreto 15x20x40cm", "unidad": "M2", "cantidad": 200},
    {"id": "C03", "partida": "TEST", "concepto": "Concepto que no existe en el catalogo", "unidad": "M3", "cantidad": 50},
]

resultados = m.emparejar(entrada)

print("\n=== MATCHES ===")
for r in resultados:
    print(f"  {r['id_entrada']} | {r['estado']:11s} | {r['metodo']:7s} | "
          f"score={r['score']} | clave={r['clave_sugerida']:8s} | PU=${r['pu_sugerido']}")

stats = m.stats_matcheo(resultados)
print(f"\nStats: {stats}")
print(f"Esperado: 2 exactos. Obtenido: {stats['exactos']} {'OK' if stats['exactos'] == 2 else 'FALLÓ'}")