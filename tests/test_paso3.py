"""Test rápido de lector_entrada e indexador."""
import json
import os
import sys
sys.path.insert(0, ".")

from src.lector_entrada import generar_entrada_ejemplo, leer_entrada, validar_entrada
from src.indexador import Indexador

# 1. Generar entrada de ejemplo
print("=" * 50)
print("1. Generando entrada de ejemplo...")
generar_entrada_ejemplo("tests/entrada_test.xlsx", 10)

# 2. Leer entrada
print("\n2. Leyendo entrada...")
conceptos = leer_entrada("tests/entrada_test.xlsx")
print(f"   Leídos: {len(conceptos)} conceptos")
for c in conceptos[:3]:
    print(f"   {c['id']} | {c['partida']:15s} | {c['concepto'][:45]} | {c['unidad']:4s} | {c['cantidad']:.2f}")

# 3. Validar
print("\n3. Validando...")
v = validar_entrada("tests/entrada_test.xlsx")
print(f"   Válido: {v['valido']}, filas: {v['num_filas']}, formato: {v['formato']} cols")

# 4. Crear catálogo mock y probar indexador
print("\n4. Creando catálogo mock (20 conceptos)...")
mock_conceptos = []
for i in range(1, 21):
    mock_conceptos.append({
        "id": f"PONTEM-{i:04d}",
        "paquete": "EDIFICACION",
        "nivel_1": "Cimentacion" if i <= 10 else "Estructura",
        "nivel_2": "Concreto" if i % 2 == 0 else "Acero",
        "clave": f"CAT-{i:03d}",
        "descripcion": f"Concepto de prueba numero {i} - concreto armado",
        "unidad": "M3" if i % 3 == 0 else ("M2" if i % 3 == 1 else "KG"),
        "precio_unitario": round(500 + i * 75.5, 2),
        "aplicacion": "general",
    })

os.makedirs("tests", exist_ok=True)
with open("tests/catalogo_mock.json", "w", encoding="utf-8") as f:
    json.dump({"conceptos": mock_conceptos, "meta": {"version": "2026-01"}}, f, indent=2)

# 5. Indexar
print("\n5. Indexando catálogo...")
idx = Indexador("tests/catalogo_mock.json")
total = idx.cargar()
print(f"   Cargados: {total} conceptos")
print(f"   Stats: {idx.stats()}")

# 6. Búsquedas
print("\n6. Búsquedas...")
r = idx.por_clave("CAT-005")
print(f"   por_clave('CAT-005'): {r['descripcion']} | ${r['precio_unitario']}")

r = idx.buscar("concreto", limite=3)
print(f"   buscar('concreto'): {len(r)} resultados")
for c in r:
    print(f"     {c['clave']} | {c['descripcion']} | ${c['precio_unitario']}")

print("\n" + "=" * 50)
print("TODO OK - Paso 3 funcionando")
print("=" * 50)