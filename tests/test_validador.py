"""
test_validador.py
==================
Pruebas del validador de integridad del libro.

El validador debe DETECTAR los defectos, no solo dar luz verde: aquí se
rompen el libro a propósito y se comprueba que los detecta.

Ejecutar:  python tests\\test_validador.py
"""

import shutil
import sys
import tempfile
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


def _libro_limpio(tmp):
    ruta, _, _ = construir_de_prueba(tmp)
    return ruta


def test_libro_limpio_pasa():
    print("\n1. Un libro íntegro no reporta errores")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = _libro_limpio(tmp)
        from src.validador import validar_libro
        rep = validar_libro(str(ruta), recalcular=False)
        check(rep["ok"], f"0 errores (obtenidos: {rep['errores'][:3]})")


def test_detecta_hoja_inexistente():
    print("\n2. Detecta una fórmula que referencia una hoja inexistente")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = _libro_limpio(tmp)
        wb = openpyxl.load_workbook(ruta)
        wb["Presupuesto"]["Z1"] = "=HojaQueNoExiste!A1"
        wb.save(ruta)
        wb.close()
        from src.validador import validar_libro
        rep = validar_libro(str(ruta), recalcular=False)
        check(not rep["ok"], "se detectaron errores")
        check(any("HojaQueNoExiste" in e for e in rep["errores"]),
              f"el error menciona la hoja inexistente ({rep['errores'][:2]})")


def test_detecta_clave_inexistente():
    print("\n3. Detecta un insumo de APU que no está en Insumos")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = _libro_limpio(tmp)
        wb = openpyxl.load_workbook(ruta)
        ws = wb["CIM-02"]
        for r in range(5, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if isinstance(v, str) and v.startswith("INS-"):
                ws.cell(row=r, column=1).value = "INS-999"
                break
        wb.save(ruta)
        wb.close()
        from src.validador import validar_libro
        rep = validar_libro(str(ruta), recalcular=False)
        check(not rep["ok"], "se detectaron errores")
        check(any("INS-999" in e for e in rep["errores"]),
              f"el error menciona la clave inexistente ({rep['errores'][:2]})")


def test_detecta_id_local_duplicado():
    print("\n3b. Detecta un id local duplicado en Insumos")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = _libro_limpio(tmp)
        wb = openpyxl.load_workbook(ruta)
        ws = wb["Insumos"]
        ids = [ws.cell(row=r, column=1).value for r in range(5, ws.max_row + 1)
               if isinstance(ws.cell(row=r, column=1).value, str)]
        ws.cell(row=6, column=1).value = ids[0]      # duplica el primero
        wb.save(ruta)
        wb.close()
        from src.validador import validar_libro
        rep = validar_libro(str(ruta), recalcular=False)
        check(not rep["ok"], "se detectaron errores")
        check(any("único" in e.lower() or "duplic" in e.lower() for e in rep["errores"]),
              f"el error es de id duplicado ({rep['errores'][:2]})")


def test_detecta_unidad_inconsistente():
    print("\n4. Detecta una unidad inconsistente entre APU y Presupuesto")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = _libro_limpio(tmp)
        wb = openpyxl.load_workbook(ruta)
        ws = wb["Presupuesto"]
        for r in range(7, ws.max_row + 1):
            if ws.cell(row=r, column=1).value == "CIM-02":
                ws.cell(row=r, column=5).value = "M3"   # el APU de CIM-02 es M2
                break
        wb.save(ruta)
        wb.close()
        from src.validador import validar_libro
        rep = validar_libro(str(ruta), recalcular=False)
        check(not rep["ok"], "se detectaron errores")
        check(any("unidad" in e.lower() for e in rep["errores"]),
              f"el error es de unidad ({rep['errores'][:2]})")


def test_detecta_cadena_de_totales_rota():
    print("\n5. Detecta la cadena de totales rota")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = _libro_limpio(tmp)
        wb = openpyxl.load_workbook(ruta)
        ws = wb["Presupuesto"]
        for r in range(1, ws.max_row + 1):
            if str(ws.cell(row=r, column=1).value).strip() == "GRAN TOTAL (IVA INCLUIDO)":
                ws.cell(row=r, column=7).value = "=G1+G2"   # referencias wrong
                break
        wb.save(ruta)
        wb.close()
        from src.validador import validar_libro
        rep = validar_libro(str(ruta), recalcular=False)
        check(not rep["ok"], "se detectaron errores")
        check(any("totales" in e.lower() or "se esperaba" in e.lower()
                  for e in rep["errores"]),
              f"el error es de la cadena de totales ({rep['errores'][:2]})")


def test_detecta_formula_malformada():
    print("\n6. Detecta una fórmula mal formada")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = _libro_limpio(tmp)
        wb = openpyxl.load_workbook(ruta)
        wb["Presupuesto"]["Z2"] = "=SUM(A1:A2"
        wb.save(ruta)
        wb.close()
        from src.validador import validar_libro
        rep = validar_libro(str(ruta), recalcular=False)
        check(not rep["ok"], "se detectaron errores")
        check(any("paréntesis" in e.lower() for e in rep["errores"]),
              f"el error es de sintaxis ({rep['errores'][:2]})")


def test_detecta_ciclo_entre_hojas():
    print("\n7. Detecta un ciclo de referencias entre hojas")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = _libro_limpio(tmp)
        wb = openpyxl.load_workbook(ruta)
        wb["Insumos"]["Z3"] = "=Presupuesto!A1"      # Insumos -> Presupuesto
        wb["Presupuesto"]["Z3"] = "=Insumos!A1"      # Presupuesto -> Insumos
        wb.save(ruta)
        wb.close()
        from src.validador import validar_libro
        rep = validar_libro(str(ruta), recalcular=False)
        check(not rep["ok"], "se detectaron errores")
        check(any("ciclo" in e.lower() for e in rep["errores"]),
              f"el error es de ciclo ({rep['errores'][:2]})")


def test_cuadre_de_totales_con_excel():
    print("\n8. Cuadre de totales (si hay Excel disponible)")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = _libro_limpio(tmp)
        from src.validador import _recalcular_excel, validar_libro
        valores = _recalcular_excel(ruta)
        if not valores:
            print("  [SKIP] Excel no disponible: el cuadre no se puede verificar")
            return
        rep = validar_libro(str(ruta), recalcular=True)
        errores_totales = [e for e in rep["errores"] if "cuadre" in e.lower()
                           or "costo directo" in e.lower() or "total" in e.lower()]
        check(not errores_totales,
              f"los totales cuadran ({errores_totales[:2]})")


def test_detecta_tabla_portada_descuadrada():
    print("\n9. Detecta que la tabla de la Portada no sume el Costo Directo")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = _libro_limpio(tmp)
        from src.validador import _recalcular_excel, validar_libro
        if not _recalcular_excel(ruta):
            print("  [SKIP] Excel no disponible: el cuadre no se puede verificar")
            return
        wb = openpyxl.load_workbook(ruta)
        ws = wb["Portada"]
        # Se altera la cantidad de un concepto SOLO en la Portada: su tabla deja
        # de sumar el Costo Directo del Presupuesto.
        fila = None
        for r in range(1, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if isinstance(v, str) and v.startswith("CIM-"):
                fila = r
                break
        if fila is None:
            check(False, "no se encontró ningún concepto en la Portada")
            wb.close()
            return
        ws.cell(row=fila, column=4).value = 999999
        wb.save(ruta)
        wb.close()
        rep = validar_libro(str(ruta), recalcular=True)
        check(not rep["ok"], "se detectaron errores")
        check(any("portada" in e.lower() for e in rep["errores"]),
              f"el error menciona la Portada ({rep['errores'][:2]})")


def test_detecta_cuadrilla_en_cero():
    print("\n10. Detecta una cuadrilla que resuelve $0 (presupuesto subvaluado)")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = _libro_limpio(tmp)
        from src.validador import _recalcular_excel, validar_libro
        if not _recalcular_excel(ruta):
            print("  [SKIP] Excel no disponible: el cuadre no se puede verificar")
            return
        wb = openpyxl.load_workbook(ruta)
        ws = wb["Cuadrillas"]
        # Se deja el $/Jornal de la cuadrilla local en 0 (como cuando la
        # clave no existe y el IFERROR del VLOOKUP devuelve 0).
        for r in range(5, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if isinstance(v, str) and v.startswith("INS-"):
                ws.cell(row=r, column=9).value = 0
                break
        wb.save(ruta)
        wb.close()
        rep = validar_libro(str(ruta), recalcular=True)
        check(not rep["ok"], "se detectaron errores")
        check(any("cuadrilla" in e.lower() for e in rep["errores"]),
              f"el error menciona la cuadrilla en $0 ({rep['errores'][:2]})")


def test_avisa_pu_distinto_del_oficial():
    print("\n11. Avisa cuando el PU cotizado difiere del oficial del catálogo")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = _libro_limpio(tmp)
        import src.validador as V

        # Simular la BD canónica: el catálogo oficial de CAT-002 es 500 y el
        # libro cotiza 350.
        original = V._pu_oficiales
        V._pu_oficiales = lambda: {"CAT-002": (500.0, "publicado")}
        try:
            rep = V.validar_libro(str(ruta), recalcular=False)
        finally:
            V._pu_oficiales = original

        check(rep["ok"], "no es un error: el libro es íntegro")
        avisos = rep["advertencias"]
        check(any("PU cotizado" in a for a in avisos),
              f"se avisó del PU divergente ({avisos[:2]})")
        check(any("$500.00" in a and "$350.00" in a for a in avisos),
              f"el aviso muestra ambos precios ({avisos[:2]})")
        check(any("impacto" in a and "30,000" in a for a in avisos),
              f"el aviso quantify el impacto en dinero ({avisos[:2]})")


def test_no_avisa_si_pu_coincide():
    print("\n12. No avisa cuando el PU cotizado coincide con el oficial")
    with tempfile.TemporaryDirectory() as tmp:
        ruta = _libro_limpio(tmp)
        import src.validador as V

        original = V._pu_oficiales
        V._pu_oficiales = lambda: {"CAT-002": (350.0, "publicado")}
        try:
            rep = V.validar_libro(str(ruta), recalcular=False)
        finally:
            V._pu_oficiales = original
        check(not any("PU cotizado" in a for a in rep["advertencias"]),
              f"no hay aviso de divergencia ({rep['advertencias'][:2]})")


def main():
    print("=== PRUEBAS DEL VALIDADOR ===")
    test_libro_limpio_pasa()
    test_detecta_hoja_inexistente()
    test_detecta_clave_inexistente()
    test_detecta_id_local_duplicado()
    test_detecta_unidad_inconsistente()
    test_detecta_cadena_de_totales_rota()
    test_detecta_formula_malformada()
    test_detecta_ciclo_entre_hojas()
    test_cuadre_de_totales_con_excel()
    test_detecta_tabla_portada_descuadrada()
    test_detecta_cuadrilla_en_cero()
    test_avisa_pu_distinto_del_oficial()
    test_no_avisa_si_pu_coincide()

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
