"""
test_matcher.py
================
Pruebas del emparejador de conceptos.

Ejecutar:  python tests\\test_matcher.py
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from tests.fixtures import ClienteFake, matcher_de_prueba  # noqa: E402

FALLOS = []


def check(condicion, mensaje):
    if condicion:
        print(f"  [OK]   {mensaje}")
    else:
        print(f"  [FALLA] {mensaje}")
        FALLOS.append(mensaje)


def test_match_exacto_misma_unidad():
    print("\n1. Match exacto con la MISMA unidad")
    m = matcher_de_prueba(umbral=55, piso=40)
    r = m.emparejar([{"id": "X1", "concepto": "Trazo y nivelacion de terreno",
                      "unidad": "M2", "cantidad": 10}])[0]
    check(r["metodo"] == "exacto", f"metodo == exacto (fue {r['metodo']})")
    check(r["estado"] == "confirmado", f"estado == confirmado (fue {r['estado']})")
    check(r["score"] == 100.0, f"score == 100 (fue {r['score']})")
    check(r["clave_sugerida"] == "CAT-001", f"clave == CAT-001 (fue {r['clave_sugerida']})")
    check(r["pu_sugerido"] == 500.0, f"PU == 500 (fue {r['pu_sugerido']})")


def test_match_exacto_ignora_unidad_ajena():
    print("\n2. Descripción idéntica en OTRA unidad no debe traer su precio")
    m = matcher_de_prueba(umbral=55, piso=40)
    # CAT-003 (KG, $45) y CAT-004 (TON, $45,000) tienen la misma descripción
    r_kg = m.emparejar([{"id": "X2", "concepto": "Acero de refuerzo fy=4200 kg/cm2",
                         "unidad": "KG", "cantidad": 10}])[0]
    r_ton = m.emparejar([{"id": "X3", "concepto": "Acero de refuerzo fy=4200 kg/cm2",
                          "unidad": "TON", "cantidad": 10}])[0]
    check(r_kg["clave_sugerida"] == "CAT-003",
          f"KG -> CAT-003 (fue {r_kg['clave_sugerida']} / PU {r_kg['pu_sugerido']})")
    check(abs(r_kg["pu_sugerido"] - 45.0) < 0.01,
          f"PU de KG == 45 (fue {r_kg['pu_sugerido']})")
    check(r_ton["clave_sugerida"] == "CAT-004",
          f"TON -> CAT-004 (fue {r_ton['clave_sugerida']} / PU {r_ton['pu_sugerido']})")
    check(abs(r_ton["pu_sugerido"] - 45000.0) < 0.01,
          f"PU de TON == 45000 (fue {r_ton['pu_sugerido']})")
    check(r_kg["pu_sugerido"] != r_ton["pu_sugerido"],
          "los precios de KG y TON son distintos")


def test_unidad_inexistente_no_toma_precio_ajeno():
    print("\n3. Unidad inexistente no toma el precio de otra unidad")
    m = matcher_de_prueba(umbral=55, piso=40)
    r = m.emparejar([{"id": "X4", "concepto": "Acero de refuerzo fy=4200 kg/cm2",
                      "unidad": "M3", "cantidad": 10}])[0]
    # La entrada pide M3 y el catálogo solo tiene esa descripción en KG y TON:
    # no se debe confirmar con el precio de TON ($45,000).
    check(r["pu_sugerido"] != 45000.0,
          f"no se aplicó el PU de TON (fue {r['pu_sugerido']})")
    check(r["estado"] != "confirmado",
          f"el resultado no queda confirmado (estado: {r['estado']})")
    check(bool(r.get("aviso")), f"se informa el cambio de unidad: '{r.get('aviso', '')[:60]}'")


def test_fuente_oficial_participa():
    print("\n4. La fuente oficial (ClienteCosteo.buscar) participa en la selección")
    cliente = ClienteFake()
    m = matcher_de_prueba(cliente=cliente, umbral=55, piso=40)
    check(len(m._clave_a_idx) == len(cliente.todos()),
          f"_clave_a_idx tiene {len(m._clave_a_idx)} claves "
          f"(esperado {len(cliente.todos())})")
    # Descripción corta: no es match exacto, pero el buscador oficial la encuentra
    r = m.emparejar([{"id": "X5", "concepto": "Muro de block", "unidad": "M2",
                      "cantidad": 10}])[0]
    check(r["metodo"] == "oficial",
          f"metodo == oficial (fue {r['metodo']})")
    check(r["clave_sugerida"] == "CAT-002", f"clave == CAT-002 (fue {r['clave_sugerida']})")


def test_normalizacion_unidades():
    print("\n5. Normalización de unidades")
    from src.matcher import normalizar_unidad
    check(normalizar_unidad("m²") == "M2", "m² -> M2")
    check(normalizar_unidad("M 2") == "M2", "'M 2' -> M2")
    check(normalizar_unidad("PZA.") == "PZA", "'PZA.' -> PZA")
    check(normalizar_unidad("metro cuadrado") == "M2", "'metro cuadrado' -> M2")
    check(normalizar_unidad("metro lineal") == "ML", "'metro lineal' -> ML")
    check(normalizar_unidad("kg") == "KG", "kg -> KG")


def test_stats_incluye_oficiales():
    print("\n6. Las estadísticas cuentan los matches oficiales")
    m = matcher_de_prueba(umbral=55, piso=40)
    res = m.emparejar([
        {"id": "A", "concepto": "Trazo y nivelacion de terreno", "unidad": "M2", "cantidad": 1},
        {"id": "B", "concepto": "Muro de block", "unidad": "M2", "cantidad": 1},
    ])
    stats = m.stats_matcheo(res)
    check(stats["exactos"] == 1, f"exactos == 1 (fue {stats['exactos']})")
    check(stats["oficiales"] == 1, f"oficiales == 1 (fue {stats['oficiales']})")
    check(stats["llm"] == 0, f"llm == 0 (fue {stats['llm']})")
    check(stats["exactos"] + stats["oficiales"] + stats["fuzzy"] + stats["llm"]
          + stats["sin_match"] == stats["total"], "los métodos suman el total")


def main():
    print("=== PRUEBAS DE MATCHER ===")
    test_match_exacto_misma_unidad()
    test_match_exacto_ignora_unidad_ajena()
    test_unidad_inexistente_no_toma_precio_ajeno()
    test_fuente_oficial_participa()
    test_normalizacion_unidades()
    test_stats_incluye_oficiales()

    print("\n" + "=" * 50)
    if FALLOS:
        print(f"FALLARON {len(FALLOS)} pruebas:")
        for f in FALLOS:
            print(f"  - {f}")
        return 1
    print("TODAS LAS PRUEBAS PASARON")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
