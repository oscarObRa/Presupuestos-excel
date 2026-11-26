"""
test_constructor.py
===================
Pruebas de construcción del libro (.xlsx) y de la hoja Portada.

Ejecutar:  python tests\\test_constructor.py
"""

import sys
import tempfile
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import openpyxl  # noqa: E402

from tests.fixtures import construir_de_prueba  # noqa: E402

FALLOS = []


def check(condicion, mensaje):
    if condicion:
        print(f"  [OK]   {mensaje}")
    else:
        print(f"  [FALLA] {mensaje}")
        FALLOS.append(mensaje)


def _filas_por_id(ws, prefijo="CIM-"):
    filas = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and v.startswith(prefijo):
            filas.append((r, v))
    return filas


def test_portada_sin_duplicados():
    print("\n1. La Portada no duplica conceptos")
    with tempfile.TemporaryDirectory() as tmp:
        ruta, _, _ = construir_de_prueba(tmp)
        wb = openpyxl.load_workbook(ruta)
        ws = wb["Portada"]
        filas = _filas_por_id(ws)
        conteo = Counter(v for _, v in filas)
        dups = {k: v for k, v in conteo.items() if v > 1}
        check(not dups, f"cada concepto aparece una vez (duplicados: {dups})")
        check(len(filas) == 3, f"la Portada tiene 3 conceptos (tiene {len(filas)})")
        wb.close()


def test_portada_sin_pie_de_totales():
    print("\n2. La Portada no copia el pie de totales del Presupuesto")
    with tempfile.TemporaryDirectory() as tmp:
        ruta, _, _ = construir_de_prueba(tmp)
        wb = openpyxl.load_workbook(ruta)
        ws = wb["Portada"]
        etiquetas = []
        for r in range(1, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if isinstance(v, str):
                etiquetas.append(v.strip())
        for prohibido in ("SUBTOTAL GENERAL (Costo Directo)", "Indirectos (% del CD)",
                          "Utilidad (% del CD + CI)", "TOTAL SIN IVA", "IVA",
                          "GRAN TOTAL (IVA INCLUIDO)"):
            # Solo se prohíbe dentro de la TABLA (antes del bloque financiero)
            check(prohibido not in etiquetas[:etiquetas.index("Costo Directo")]
                  if "Costo Directo" in etiquetas else True,
                  f"'{prohibido}' no aparece como concepto en la tabla")
        # El bloque financiero propio sí existe y refers al Presupuesto
        fila_total = next((r for r in range(1, ws.max_row + 1)
                           if str(ws.cell(row=r, column=1).value).strip() == "Total"), None)
        check(fila_total is not None, "la Portada tiene su fila 'Total'")
        if fila_total:
            formula = ws.cell(row=fila_total, column=6).value
            check(isinstance(formula, str) and formula.startswith("=Presupuesto!"),
                  f"el Total de la Portada referencia al Presupuesto (fue {formula!r})")
        wb.close()


def test_portada_importes_en_su_propia_fila():
    print("\n3. Los importes de la Portada se calculan en su propia fila")
    with tempfile.TemporaryDirectory() as tmp:
        ruta, _, _ = construir_de_prueba(tmp)
        wb = openpyxl.load_workbook(ruta)
        ws = wb["Portada"]
        filas = _filas_por_id(ws)
        ok = True
        detalle = []
        for r, _ in filas:
            formula = ws.cell(row=r, column=6).value
            # En la Portada el P.U. está en la columna E (en el Presupuesto está
            # en F): "=D*F" sería una referencia circular.
            esperado = f"=D{r}*E{r}"
            if formula != esperado:
                ok = False
                detalle.append(f"fila {r}: {formula!r} != {esperado!r}")
        check(ok, f"importes con la fórmula de su propia fila {detalle}")
        wb.close()


def test_presupuesto_sin_duplicados():
    print("\n4. La hoja Presupuesto lista cada concepto una vez")
    with tempfile.TemporaryDirectory() as tmp:
        ruta, _, _ = construir_de_prueba(tmp)
        wb = openpyxl.load_workbook(ruta)
        ws = wb["Presupuesto"]
        filas = _filas_por_id(ws)
        conteo = Counter(v for _, v in filas)
        dups = {k: v for k, v in conteo.items() if v > 1}
        check(not dups, f"cada concepto aparece una vez (duplicados: {dups})")
        wb.close()


def test_hojas_apu_por_id_entrada():
    print("\n5. Las hojas APU se nombran con el id de entrada")
    with tempfile.TemporaryDirectory() as tmp:
        ruta, _, _ = construir_de_prueba(tmp)
        wb = openpyxl.load_workbook(ruta)
        check("CIM-02" in wb.sheetnames, f"existe la hoja APU 'CIM-02' ({wb.sheetnames})")
        check("Revision" not in wb.sheetnames,
              "no se genera la hoja Revision (incluir_revision=False)")
        wb.close()


def test_validacion_sin_errores():
    print("\n6. El libro construido pasa la validación")
    with tempfile.TemporaryDirectory() as tmp:
        ruta, _, _ = construir_de_prueba(tmp)
        from src.validador import validar_libro
        reporte = validar_libro(str(ruta), recalcular=False)
        check(reporte["ok"], f"0 errores de validación ({reporte['errores'][:3]})")
        for c in reporte["checks"]:
            print(f"         [{'OK ' if c['ok'] else 'ERR'}] {c['nombre']}"
                  + (f" -> {c['detalle']}" if c["detalle"] else ""))


def test_ids_locales_en_insumos_y_apu():
    print("\n7. Los insumos copiados usan IDs locales y las APU los usan")
    with tempfile.TemporaryDirectory() as tmp:
        ruta, _, _ = construir_de_prueba(tmp)
        wb = openpyxl.load_workbook(ruta)
        ws = wb["Insumos"]
        ids = []
        origenes = []
        for r in range(5, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if isinstance(v, str) and v.strip():
                ids.append(v.strip())
                origenes.append(ws.cell(row=r, column=9).value)
        check(all(i.startswith("INS-") for i in ids),
              f"todos los ids de Insumos son locales (INS-xxx): {ids}")
        check(len(set(ids)) == len(ids), f"los ids locales son únicos: {ids}")
        check(any(o == "1A1P" for o in origenes),
              f"la clave general se conserva en 'Clave origen': {origenes}")

        # Las líneas de las hojas APU deben usar el MISMO id local
        ws_apu = wb["CIM-02"]
        ids_apu = []
        for r in range(5, ws_apu.max_row + 1):
            v = ws_apu.cell(row=r, column=1).value
            if isinstance(v, str) and v.startswith("INS-"):
                ids_apu.append(v.strip())
        check(bool(ids_apu) and set(ids_apu).issubset(set(ids)),
              f"las APU usan ids locales de Insumos: {ids_apu} vs {ids}")
        wb.close()


def test_cuadrillas_registradas_con_id_local():
    print("\n8. Las cuadrillas copiadas se registran en la hoja Cuadrillas")
    with tempfile.TemporaryDirectory() as tmp:
        ruta, _, _ = construir_de_prueba(tmp)
        wb = openpyxl.load_workbook(ruta)
        ws_ins = wb["Insumos"]
        ws_cuad = wb["Cuadrillas"]
        cuadrillas = []
        for r in range(5, ws_ins.max_row + 1):
            if str(ws_ins.cell(row=r, column=2).value or "").strip().lower() == "cuadrilla":
                cuadrillas.append((ws_ins.cell(row=r, column=1).value,
                                   ws_ins.cell(row=r, column=5).value,
                                   r))
        check(bool(cuadrillas), f"hay cuadrillas en Insumos: {cuadrillas}")

        ids_cuad = {}
        for r in range(5, ws_cuad.max_row + 1):
            v = ws_cuad.cell(row=r, column=1).value
            if isinstance(v, str) and v.strip():
                ids_cuad[v.strip()] = r
        for id_local, formula, fila_ins in cuadrillas:
            check(id_local in ids_cuad,
                  f"la cuadrilla {id_local} existe en la hoja Cuadrillas")
            if id_local in ids_cuad:
                fila = ids_cuad[id_local]
                tasa = ws_cuad.cell(row=fila, column=9).value
                check(isinstance(tasa, (int, float)) and tasa > 0,
                      f"la cuadrilla {id_local} tiene $/Jornal ({tasa})")
                # El VLOOKUP debe apuntar a la celda del id local (A{fila})
                check(isinstance(formula, str) and f"A{fila_ins}" in formula,
                      f"el VLOOKUP de {id_local} usa el id local de su fila "
                      f"(A{fila_ins}): {formula}")
        wb.close()


def test_pu_oficial_se_cotiza():
    print("\n9. El PU cotizado es el oficial del modelo canónico")
    with tempfile.TemporaryDirectory() as tmp:
        ruta, _, _ = construir_de_prueba(tmp)
        wb = openpyxl.load_workbook(ruta)
        ws = wb["Presupuesto"]
        # CAT-002: PU oficial 350, PU del APU 372.5 -> debe cotizar 350
        fila = None
        for r, _id in _filas_por_id(ws):
            if ws.cell(row=r, column=2).value == "CAT-002":
                fila = r
                break
        check(fila is not None, "el concepto CAT-002 está en el Presupuesto")
        if fila:
            pu = ws.cell(row=fila, column=6).value
            check(isinstance(pu, (int, float)), f"el PU es un valor (no fórmula): {pu!r}")
            check(abs(float(pu) - 350.0) < 0.01,
                  f"el PU cotizado es el oficial (350), no el del APU (372.5): {pu}")
        wb.close()


def test_modo_cotizar_apu():
    print("\n10. Con usar_pu_oficial=False se cotiza el PU de la hoja APU")
    with tempfile.TemporaryDirectory() as tmp:
        ruta, _, _ = construir_de_prueba(tmp, usar_pu_oficial=False)
        wb = openpyxl.load_workbook(ruta)
        ws = wb["Presupuesto"]
        fila = None
        for r, _id in _filas_por_id(ws):
            if ws.cell(row=r, column=2).value == "CAT-002":
                fila = r
                break
        check(fila is not None, "el concepto CAT-002 está en el Presupuesto")
        if fila:
            pu = ws.cell(row=fila, column=6).value
            check(isinstance(pu, str) and pu.startswith("="),
                  f"el PU es una referencia a la hoja APU: {pu!r}")
            check("'CIM-02'" in pu or "CIM-02" in pu,
                  f"la referencia apunta a la hoja APU del concepto: {pu}")
        wb.close()


def test_revision_muestra_trazabilidad_del_precio():
    print("\n11. La Revisión muestra PU oficial, PU APU, diferencia, origen y fuente")
    with tempfile.TemporaryDirectory() as tmp:
        ruta, _, _ = construir_de_prueba(tmp, incluir_revision=True)
        wb = openpyxl.load_workbook(ruta)
        ws = wb["Revision"]
        headers = [ws.cell(row=4, column=c).value for c in range(1, ws.max_column + 1)]
        for h in ("PU oficial", "PU APU", "Dif. %", "Origen precio", "Fuente desglose"):
            check(h in headers, f"existe la columna '{h}': {headers}")

        fila = None
        for r in range(5, ws.max_row + 1):
            if ws.cell(row=r, column=5).value == "CAT-002":
                fila = r
                break
        check(fila is not None, "CAT-002 aparece en la Revisión")
        if fila:
            pu_of = ws.cell(row=fila, column=11).value
            pu_apu = ws.cell(row=fila, column=12).value
            dif = ws.cell(row=fila, column=13).value
            origen = ws.cell(row=fila, column=14).value
            fuente = ws.cell(row=fila, column=15).value
            check(abs(float(pu_of) - 350.0) < 0.01, f"PU oficial = 350: {pu_of}")
            check(abs(float(pu_apu) - 372.5) < 0.01, f"PU APU = 372.5: {pu_apu}")
            check(abs(float(dif) - (372.5 - 350.0) / 350.0) < 1e-6,
                  f"la diferencia % está calculada: {dif}")
            check(origen == "publicado", f"el origen del precio es 'publicado': {origen}")
            check(fuente == "canonica", f"la fuente del desglose es 'canonica': {fuente}")
        wb.close()


def test_hoja_apu_muestra_pu_oficial():
    print("\n12. La hoja APU muestra el PU oficial junto al PU de venta")
    with tempfile.TemporaryDirectory() as tmp:
        ruta, _, _ = construir_de_prueba(tmp)
        wb = openpyxl.load_workbook(ruta)
        ws = wb["CIM-02"]
        etiquetas = {}
        for r in range(1, ws.max_row + 1):
            for c in (3, 4, 5):
                v = ws.cell(row=r, column=c).value
                if isinstance(v, str) and v.strip() in ("PU DE VENTA", "PU OFICIAL (CATALOGO)"):
                    etiquetas[v.strip()] = r
        check("PU DE VENTA" in etiquetas, "la hoja APU tiene 'PU DE VENTA'")
        check("PU OFICIAL (CATALOGO)" in etiquetas,
              f"la hoja APU tiene 'PU OFICIAL (CATALOGO)': {etiquetas}")
        if "PU OFICIAL (CATALOGO)" in etiquetas:
            r = etiquetas["PU OFICIAL (CATALOGO)"]
            pu = ws.cell(row=r, column=6).value
            check(abs(float(pu) - 350.0) < 0.01, f"el PU oficial vale 350: {pu}")
        wb.close()


def main():
    print("=== PRUEBAS DE CONSTRUCTOR ===")
    test_portada_sin_duplicados()
    test_portada_sin_pie_de_totales()
    test_portada_importes_en_su_propia_fila()
    test_presupuesto_sin_duplicados()
    test_hojas_apu_por_id_entrada()
    test_validacion_sin_errores()
    test_ids_locales_en_insumos_y_apu()
    test_cuadrillas_registradas_con_id_local()
    test_pu_oficial_se_cotiza()
    test_modo_cotizar_apu()
    test_revision_muestra_trazabilidad_del_precio()
    test_hoja_apu_muestra_pu_oficial()

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
