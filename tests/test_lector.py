# -*- coding: utf-8 -*-
"""
test_lector.py
==============
Pruebas del lector de archivos de entrada (src/lector_entrada.py).

Casos cubiertos:
1. Formato de 4 columnas y de 5 columnas.
2. Partidas/subpartidas incrustadas.
3. Columnas 'Unnamed: N' VACÍAS (celdas con formato sobrante) se descartan
   en vez de rechazar el archivo (bug real reportado en la web).
4. Columnas 'Unnamed: N' CON datos → se rechaza (datos fuera de formato).

Ejecutar:  python tests/test_lector.py
"""
import sys
import tempfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Border

from src.lector_entrada import leer_entrada, validar_entrada


def _archivo(df: pd.DataFrame, carpeta: Path, nombre: str) -> str:
    ruta = carpeta / nombre
    df.to_excel(ruta, index=False)
    return str(ruta)


def test_formatos_basicos(carpeta: Path):
    df4 = pd.DataFrame({
        "id": ["A-1", "A-2"], "concepto": ["Losa", "Muro"],
        "unidad": ["m2", "m2"], "cantidad": [10.0, 5.0],
    })
    cs = leer_entrada(_archivo(df4, carpeta, "f4.xlsx"))
    assert len(cs) == 2 and cs[0]["id"] == "A-1"

    df5 = pd.DataFrame({
        "id": ["B-1"], "partida": ["MUROS"], "concepto": ["Muro block"],
        "unidad": ["m2"], "cantidad": [5.0],
    })
    cs = leer_entrada(_archivo(df5, carpeta, "f5.xlsx"))
    assert cs[0]["partida"] == "MUROS"
    print("  formatos 4 y 5 columnas OK")


def test_partidas_incrustadas(carpeta: Path):
    df = pd.DataFrame({
        "id": ["PARTIDA 01: CIMENTACION", "CIM-01", "SUBPARTIDA 1.1 ZANJAS", "CIM-02"],
        "concepto": [None, "Excavación", None, "Relleno"],
        "unidad": [None, "m3", None, "m3"],
        "cantidad": [None, 9.75, None, 0.5],
    })
    cs = leer_entrada(_archivo(df, carpeta, "partidas.xlsx"))
    assert len(cs) == 2
    assert cs[0]["partida"].startswith("PARTIDA 01")
    assert cs[1]["subpartida"].startswith("SUBPARTIDA")
    print("  partidas/subpartidas incrustadas OK")


def test_columnas_fantasma_vacias(carpeta: Path):
    """El bug real: Excel con celdas con formato -> 'Unnamed: 4/5' vacías."""
    wb = Workbook()
    ws = wb.active
    ws.append(["id", "concepto", "unidad", "cantidad"])
    ws.append(["CIM-01", "Excavación en cepa", "m3", 9.75])
    ws.append(["CIM-02", "Relleno compactado", "m3", 0.5])
    # Celdas con formato (borde) pero sin datos → pandas las lee como Unnamed
    borde = Border(right=None)
    for r in range(1, 4):
        for c in range(5, 8):
            ws.cell(row=r, column=c).border = borde
    ws.append([None] * 7)  # fila con celdas vacías en las columnas fantasma
    ruta = carpeta / "fantasma.xlsx"
    wb.save(ruta)

    cs = leer_entrada(str(ruta))
    assert len(cs) == 2, f"esperaba 2 conceptos, hay {len(cs)}"
    assert cs[0]["id"] == "CIM-01" and cs[0]["cantidad"] == 9.75
    print("  columnas 'Unnamed' vacías descartadas OK")


def test_unnamed_con_datos_se_rechaza(carpeta: Path):
    """Una columna extra CON datos no se descarta en silencio: se rechaza."""
    df = pd.DataFrame({
        "id": ["A-1"], "concepto": ["Losa"], "unidad": ["m2"],
        "cantidad": [10.0], "extra": ["dato corrupto"],
    })
    res = validar_entrada(_archivo(df, carpeta, "extra.xlsx"))
    assert not res["valido"] and "Columnas inesperadas" in (res["error"] or "")
    print("  columna extra con datos rechazada OK")


def main() -> int:
    print("=" * 60)
    print("SUITE test_lector.py — lector de archivos de entrada")
    print("=" * 60)
    with tempfile.TemporaryDirectory() as tmp:
        carpeta = Path(tmp)
        test_formatos_basicos(carpeta)
        test_partidas_incrustadas(carpeta)
        test_columnas_fantasma_vacias(carpeta)
        test_unnamed_con_datos_se_rechaza(carpeta)
    print("\ntest_lector.py: TODAS LAS PRUEBAS PASARON")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
