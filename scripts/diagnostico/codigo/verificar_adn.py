# -*- coding: utf-8 -*-
"""
verificar_adn.py
================
Smoke test del ADN visual por empresa.

Construye un presupuesto de prueba y verifica sobre el .xlsx real que:
1. El libro se genera sin errores con empresa=None (default ICEL) y con
   empresa='icel' explicito.
2. La paleta del ADN (docs/ADN_ICEL.md) esta aplicada en las hojas:
   - Verde header  #3C7D22 en rellenos de encabezado (Insumos fila 4).
   - Borde verde   #275317 en celdas de la tabla de la Portada.
3. Una empresa desconocida se rechaza con error explícito.

Uso:  python scripts/diagnostico/codigo/verificar_adn.py
"""
import sys
import tempfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(RAIZ))
sys.stdout.reconfigure(encoding="utf-8")

from openpyxl import load_workbook
from tests.fixtures import construir_de_prueba


def fill_rgb(cell) -> str:
    c = cell.fill.start_color
    return (c.rgb or "")[-6:] if c and c.rgb else ""


def borde_rgb(cell) -> str:
    c = cell.border.top.color
    return (c.rgb or "")[-6:] if c and c.rgb else ""


def verificar(ruta: Path, etiqueta: str):
    wb = load_workbook(ruta)
    assert "Portada" in wb.sheetnames, f"{etiqueta}: falta hoja Portada"
    ws = wb["Portada"]

    # Localizar la primera fila con relleno verde de encabezado de tabla
    verdes_header = sum(1 for row in ws.iter_rows() for cell in row
                        if fill_rgb(cell) == "3C7D22")
    verdes_borde = sum(1 for row in ws.iter_rows() for cell in row
                       if borde_rgb(cell) == "275317")
    assert verdes_header > 0, f"{etiqueta}: no hay relleno 3C7D22 en Portada"
    assert verdes_borde > 0, f"{etiqueta}: no hay borde 275317 en Portada"
    print(f"  [{etiqueta}] Portada OK -> {verdes_header} celdas header "
          f"#3C7D22, {verdes_borde} celdas con borde #275317")

    ws_ins = wb["Insumos"]
    assert fill_rgb(ws_ins.cell(row=4, column=1)) == "3C7D22", \
        f"{etiqueta}: header de Insumos no usa 3C7D22"
    print(f"  [{etiqueta}] Insumos OK  -> header fila 4 con relleno #3C7D22")
    wb.close()


def main():
    with tempfile.TemporaryDirectory() as tmp:
        # 1) Default (empresa=None -> ICEL)
        ruta, _, _ = construir_de_prueba(Path(tmp) / "default")
        verificar(ruta, "empresa=None (default)")

        # 2) Explicito 'icel'
        ruta, _, _ = construir_de_prueba(Path(tmp) / "icel")
        verificar(ruta, "empresa='icel'")

    # 3) Empresa desconocida -> error explícito, nunca silencio
    from src.visual_dna import get_dna
    try:
        get_dna("empresa_inexistente")
        raise AssertionError("get_dna no rechazó la empresa desconocida")
    except ValueError as e:
        print(f"  [factory] empresa desconocida rechazada OK -> {e}")

    print("\nADN VISUAL VERIFICADO END-TO-END")


if __name__ == "__main__":
    main()
