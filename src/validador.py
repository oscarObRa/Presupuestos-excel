"""
validador.py
============
Validación de integridad del libro de presupuesto generado (.xlsx).

Ejecuta las verificaciones que anuncia el README:

 1. Toda clave usada en un VLOOKUP a Insumos existe en la hoja Insumos.
 2. Toda clave usada en las hojas APU existe en Insumos.
 2b. Los IDs locales de Insumos son únicos (la liga depende de ellos).
 3. Toda categoría usada en Cuadrillas existe en Calculo_categoria.
 4. Consistencia de unidades entre Presupuesto, APU e Insumos.
 5. Ausencia de ciclos de referencia entre hojas.
 6. Sintaxis válida de las fórmulas (paréntesis balanceados, sin #REF!).
 7. Cuadre de totales (cadena de fórmulas + recálculo con Excel si existe).
 8. Ninguna cuadrilla queda en $0 (el IFERROR del VLOOKUP lo ocultaría).
 9. El PU cotizado coincide con el PU oficial del catálogo (advertencia con el
    impacto en dinero cuando difiere; se omite si la BD canónica no está).

Uso:
    from src.validador import validar_libro

    reporte = validar_libro("salidas/presupuesto.xlsx")
    if reporte["errores"]:
        for e in reporte["errores"]:
            print("ERROR:", e)
"""

from __future__ import annotations

import re
import shutil
import tempfile
from pathlib import Path
from typing import Optional

# Hojas que no son "maestros" ni APUs (fijas de la plantilla)
HOJAS_FIJAS = {
    "portada", "presupuesto", "parametros", "calculo_categoria",
    "cuadrillas", "variables", "insumos", "notas", "categorias", "revision",
}

# Etiquetas del pie de totales de la hoja Presupuesto -> clave lógica
PIE_PRESUPUESTO = {
    "SUBTOTAL GENERAL (Costo Directo)": "cd",
    "Indirectos (% del CD)": "ind",
    "Utilidad (% del CD + CI)": "util",
    "TOTAL SIN IVA": "sin_iva",
    "IVA": "iva",
    "GRAN TOTAL (IVA INCLUIDO)": "gran",
}

VLOOKUP_RE = re.compile(r"VLOOKUP\(\s*([^,]+?)\s*,\s*([^,]+?)\s*,\s*(\d+)\s*,", re.IGNORECASE)
RANGO_RE = re.compile(
    r"^\s*(?:'([^']+)'|([^!]+?))!\$?([A-Z]{1,2})\$?(\d+):\$?([A-Z]{1,2})\$?(\d+)\s*$"
)
CELDA_RE = re.compile(r"^\$?([A-Z]{1,2})\$?(\d+)$")
REF_HOJA_RE = re.compile(r"(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_.]*))!")


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def _es_formula(val) -> bool:
    return isinstance(val, str) and val.startswith("=")


def _norm_hoja(nombre: str) -> str:
    return str(nombre or "").strip().lower()


def _referencias_hoja(formula: str) -> set[str]:
    """Nombres de hoja mencionados en una fórmula (con o sin comillas)."""
    refs = set()
    for m in REF_HOJA_RE.finditer(formula):
        nombre = (m.group(1) or m.group(2) or "").strip()
        if nombre:
            refs.add(nombre)
    return refs


def _parse_rango(expr: str) -> Optional[dict]:
    """'Insumos!$A$5:$G$20' -> {'hoja': 'Insumos', 'col': 'A', ...}"""
    m = RANGO_RE.match(expr or "")
    if not m:
        return None
    hoja = (m.group(1) or m.group(2) or "").strip()
    if not hoja:
        return None
    return {
        "hoja": hoja,
        "col": m.group(3),
        "fila_ini": int(m.group(4)),
        "col_fin": m.group(5),
        "fila_fin": int(m.group(6)),
    }


def _valor_celda(ws, coord: str):
    try:
        return ws[coord].value
    except Exception:
        return None


def _claves_maestro(wb, hoja: str, col: str, fila_ini: int, fila_fin: int) -> set[str]:
    """Claves (texto) de un rango maestro, por ejemplo Insumos!A5:A200."""
    if hoja not in wb.sheetnames:
        return set()
    ws = wb[hoja]
    claves = set()
    for r in range(fila_ini, min(fila_fin, ws.max_row) + 1):
        v = ws[f"{col}{r}"].value
        if isinstance(v, str) and v.strip():
            claves.add(v.strip())
    return claves


def _hojas_apu(wb) -> list[str]:
    return [ws.title for ws in wb.worksheets if _norm_hoja(ws.title) not in HOJAS_FIJAS]


def _filas_concepto_presupuesto(ws) -> list[int]:
    """Filas de concepto del Presupuesto (columna B = clave PONTEM)."""
    filas = []
    for r in range(7, ws.max_row + 1):
        clave = ws.cell(row=r, column=2).value
        if clave not in (None, ""):
            filas.append(r)
    return filas


def _es_fila_insumo_apu(ws, r: int) -> bool:
    """Fila de insumo de una hoja APU (columna A = clave, D = cantidad).

    Las filas '%' (unidad p.ej. '(%)MO': herramienta menor 3%, seguridad 2%)
    NO son insumos: son porcentajes sobre el subtotal de Mano de Obra que se
    calculan dentro de la propia hoja APU. No se sembraron en Insumos a
    propósito, así que los checks de claves y unidades las exentan.
    """
    clave = ws.cell(row=r, column=1).value
    cant = ws.cell(row=r, column=4).value
    if not (isinstance(clave, str) and clave.strip() != "" and isinstance(cant, (int, float))):
        return False
    unidad = str(ws.cell(row=r, column=3).value or "")
    return "%" not in unidad.upper()


def _normalizar_unidad(unidad: str) -> str:
    """Versión local (no importa src.matcher) para no acoplar el validador."""
    try:
        from src.matcher import normalizar_unidad as nu
        return nu(unidad)
    except Exception:
        return str(unidad or "").strip().upper()


# ---------------------------------------------------------------------------
# Recálculo con Excel (opcional)
# ---------------------------------------------------------------------------

def _recalcular_excel(ruta: Path) -> Optional[dict]:
    """Recalcula una COPIA del libro con Excel y devuelve valores numéricos.

    Devuelve None si no hay Excel/pywin32 disponible (entonces el cuadre de
    totales se reporta como advertencia, no como error).
    """
    try:
        import win32com.client as win32
    except Exception:
        return None

    tmp = Path(tempfile.mkdtemp(prefix="pontem_val_"))
    copia = tmp / ruta.name
    excel = None
    wb = None
    try:
        shutil.copy2(ruta, copia)
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(copia))
        try:
            excel.CalculateFull()
        except Exception:
            pass
        wb.Save()
        wb.Close(SaveChanges=False)
        wb = None
        excel.Quit()
        excel = None

        import openpyxl
        wbv = openpyxl.load_workbook(copia, data_only=True)
        valores = {}
        for ws in wbv.worksheets:
            vals = {}
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, (int, float)):
                        vals[c.coordinate] = float(c.value)
            valores[ws.title] = vals
        wbv.close()
        return valores
    except Exception:
        return None
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        shutil.rmtree(tmp, ignore_errors=True)


# ---------------------------------------------------------------------------
# Verificaciones
# ---------------------------------------------------------------------------

def _check_vlookups(wb) -> list[str]:
    """Check 1: toda clave de VLOOKUP existe en la hoja destino."""
    errores = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                formula = cell.value
                if not _es_formula(formula):
                    continue
                for m in VLOOKUP_RE.finditer(formula):
                    key_expr, rango_expr = m.group(1), m.group(2)
                    rango = _parse_rango(rango_expr)
                    if not rango:
                        continue
                    destino = rango["hoja"]
                    if destino not in wb.sheetnames:
                        continue
                    # Resolver la clave buscada
                    clave = None
                    mc = CELDA_RE.match(key_expr.strip())
                    if mc:
                        clave = _valor_celda(ws, f"{mc.group(1)}{mc.group(2)}")
                    elif key_expr.strip().startswith('"'):
                        clave = key_expr.strip().strip('"')
                    if not isinstance(clave, str) or not clave.strip():
                        continue
                    clave = clave.strip()
                    claves = _claves_maestro(wb, destino, rango["col"],
                                             rango["fila_ini"], rango["fila_fin"])
                    if clave not in claves:
                        errores.append(
                            f"{ws.title}!{cell.coordinate}: VLOOKUP de la clave "
                            f"'{clave}' no existe en {destino}!{rango['col']}"
                            f"{rango['fila_ini']}:{rango['col']}{rango['fila_fin']}"
                        )
    return errores


def _check_apu_claves(wb) -> list[str]:
    """Check 2: toda clave de insumo de las hojas APU existe en Insumos."""
    if "Insumos" not in wb.sheetnames:
        return []
    claves_insumos = _claves_maestro(wb, "Insumos", "A", 5, 5000)
    errores = []
    for nombre in _hojas_apu(wb):
        ws = wb[nombre]
        for r in range(5, ws.max_row + 1):
            if not _es_fila_insumo_apu(ws, r):
                continue
            clave = str(ws.cell(row=r, column=1).value).strip()
            if clave not in claves_insumos:
                errores.append(
                    f"{nombre}!A{r}: el insumo '{clave}' no existe en la hoja Insumos"
                )
    return errores


def _check_ids_locales(wb) -> list[str]:
    """Check 2b: los IDs locales de Insumos son únicos (la liga depende de ellos).

    Con IDs duplicados el VLOOKUP devolvería la primera coincidencia y las
    hojas APU tomarían el costo de otro insumo.
    """
    if "Insumos" not in wb.sheetnames:
        return []
    ws = wb["Insumos"]
    vistos: dict[str, int] = {}
    errores = []
    for r in range(5, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if not isinstance(v, str) or not v.strip():
            continue
        id_local = v.strip()
        if id_local in vistos:
            errores.append(
                f"Insumos!A{r}: el id local '{id_local}' ya se usa en la fila "
                f"{vistos[id_local]} (debe ser único)"
            )
        else:
            vistos[id_local] = r
    return errores


def _check_categorias(wb) -> list[str]:
    """Check 3: toda categoría de Cuadrillas existe en Calculo_categoria."""
    if "Cuadrillas" not in wb.sheetnames or "Calculo_categoria" not in wb.sheetnames:
        return []
    ws = wb["Cuadrillas"]
    categorias = _claves_maestro(wb, "Calculo_categoria", "A", 5, 200)
    errores = []
    for r in range(5, ws.max_row + 1):
        id_cuad = ws.cell(row=r, column=1).value
        if not id_cuad:
            continue
        for col in (4, 6, 8):          # Categoría 1, 2, 3
            cat = ws.cell(row=r, column=col).value
            if not isinstance(cat, str) or not cat.strip():
                continue
            cat = cat.strip()
            if cat not in categorias:
                errores.append(
                    f"Cuadrillas!{ws.cell(row=r, column=col).coordinate}: categoría "
                    f"'{cat}' no existe en Calculo_categoria!A5:A200"
                )
    return errores


def _hoja_apu_por_id(wb, id_entrada: str) -> Optional[str]:
    """Hoja APU que corresponde a un id de entrada.

    El encabezado C1 de cada hoja APU guarda el id de entrada (ej. CIM-02),
    asi que se puede localizar aunque el PU del Presupuesto sea un valor
    (PU oficial del catalogo) y no una referencia a la hoja.
    """
    if not id_entrada:
        return None
    objetivo = str(id_entrada).strip()
    for nombre in _hojas_apu(wb):
        v = wb[nombre].cell(row=1, column=3).value
        if isinstance(v, str) and v.strip() == objetivo:
            return nombre
    return None


def _check_unidades(wb) -> list[str]:
    """Check 4: unidades consistentes entre Presupuesto, APU e Insumos."""
    errores = []
    # 4a. Presupuesto vs hoja APU referenciada
    if "Presupuesto" in wb.sheetnames:
        ws = wb["Presupuesto"]
        for r in _filas_concepto_presupuesto(ws):
            uni_pres = ws.cell(row=r, column=5).value
            pu = ws.cell(row=r, column=6).value
            hoja_apu = None
            if _es_formula(pu):
                m = re.search(r"'([^']+)'!\$?[A-Z]{1,2}\$?\d+", pu)
                if m:
                    hoja_apu = m.group(1)
            else:
                # PU oficial (valor): la hoja APU se localiza por el id de
                # entrada de la fila (columna A)
                hoja_apu = _hoja_apu_por_id(wb, ws.cell(row=r, column=1).value)
            if not hoja_apu or hoja_apu not in wb.sheetnames:
                continue
            uni_apu = wb[hoja_apu].cell(row=2, column=2).value
            if not uni_pres or not uni_apu:
                continue
            if _normalizar_unidad(uni_pres) != _normalizar_unidad(uni_apu):
                errores.append(
                    f"Presupuesto!E{r}: unidad '{uni_pres}' no coincide con la "
                    f"unidad '{uni_apu}' de la hoja APU '{hoja_apu}'"
                )
    # 4b. APU vs Insumos (misma clave)
    if "Insumos" in wb.sheetnames:
        ws_ins = wb["Insumos"]
        unidades_ins = {}
        for r in range(5, ws_ins.max_row + 1):
            k = ws_ins.cell(row=r, column=1).value
            if isinstance(k, str) and k.strip():
                unidades_ins.setdefault(k.strip(), ws_ins.cell(row=r, column=4).value)
        for nombre in _hojas_apu(wb):
            ws = wb[nombre]
            for r in range(5, ws.max_row + 1):
                if not _es_fila_insumo_apu(ws, r):
                    continue
                clave = str(ws.cell(row=r, column=1).value).strip()
                uni_apu = ws.cell(row=r, column=3).value
                uni_ins = unidades_ins.get(clave)
                if not uni_apu or not uni_ins:
                    continue
                if _normalizar_unidad(uni_apu) != _normalizar_unidad(uni_ins):
                    errores.append(
                        f"{nombre}!C{r}: unidad '{uni_apu}' del insumo '{clave}' "
                        f"no coincide con '{uni_ins}' en Insumos"
                    )
    return errores


def _check_ciclos(wb) -> list[str]:
    """Check 5: grafo de referencias entre hojas sin ciclos."""
    nombres = {ws.title for ws in wb.worksheets}
    grafo = {}
    for ws in wb.worksheets:
        destinos = set()
        for row in ws.iter_rows():
            for cell in row:
                if not _es_formula(cell.value):
                    continue
                for ref in _referencias_hoja(cell.value):
                    if ref in nombres:
                        destinos.add(ref)
        grafo[ws.title] = destinos

    ciclos = []
    estado = {}   # 0 = sin visitar, 1 = en curso, 2 = terminado
    pila = []

    def _dfs(nodo: str):
        estado[nodo] = 1
        pila.append(nodo)
        for sig in sorted(grafo.get(nodo, ())):
            if estado.get(sig, 0) == 1:
                idx = pila.index(sig)
                ciclos.append(" -> ".join(pila[idx:] + [sig]))
            elif estado.get(sig, 0) == 0:
                _dfs(sig)
        pila.pop()
        estado[nodo] = 2

    for nodo in sorted(grafo):
        if estado.get(nodo, 0) == 0:
            _dfs(nodo)
    return [f"ciclo de referencias entre hojas: {c}" for c in sorted(set(ciclos))]


def _check_formulas(wb) -> list[str]:
    """Check 6: sintaxis básica de las fórmulas y hojas existentes."""
    nombres = {ws.title for ws in wb.worksheets}
    errores = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                formula = cell.value
                if not _es_formula(formula):
                    continue
                if formula.count("(") != formula.count(")"):
                    errores.append(f"{ws.title}!{cell.coordinate}: paréntesis desbalanceados")
                if formula.count('"') % 2 != 0:
                    errores.append(f"{ws.title}!{cell.coordinate}: comillas sin cerrar")
                if "#REF!" in formula:
                    errores.append(f"{ws.title}!{cell.coordinate}: referencia #REF!")
                for ref in _referencias_hoja(formula):
                    if ref not in nombres:
                        errores.append(
                            f"{ws.title}!{cell.coordinate}: referencia a la hoja "
                            f"'{ref}' que no existe en el libro"
                        )
    return errores


def _pie_presupuesto(ws) -> dict:
    """Localiza las filas del pie de totales del Presupuesto."""
    pie = {}
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str):
            clave = PIE_PRESUPUESTO.get(v.strip())
            if clave and clave not in pie:
                pie[clave] = r
    return pie


def _check_totales_estructura(wb) -> list[str]:
    """Check 7 (estructura): la cadena de totales apunta a las filas correctas."""
    if "Presupuesto" not in wb.sheetnames:
        return []
    ws = wb["Presupuesto"]
    pie = _pie_presupuesto(ws)
    faltan = [k for k in ("cd", "ind", "util", "sin_iva", "iva", "gran") if k not in pie]
    if faltan:
        return [f"Presupuesto: faltan filas del pie de totales ({', '.join(faltan)})"]

    errores = []
    cd, ind, util, sin_iva, iva, gran = (pie["cd"], pie["ind"], pie["util"],
                                          pie["sin_iva"], pie["iva"], pie["gran"])
    esperado = {
        (ind, 7): f"=G{cd}*F{ind}",
        (util, 7): f"=(G{cd}+G{ind})*F{util}",
        (sin_iva, 7): f"=G{cd}+G{ind}+G{util}",
        (iva, 7): f"=G{sin_iva}*F{iva}",
        (gran, 7): f"=G{sin_iva}+G{iva}",
    }
    for (fila, col), formula_esperada in esperado.items():
        real = ws.cell(row=fila, column=col).value
        if real != formula_esperada:
            errores.append(
                f"Presupuesto!{ws.cell(row=fila, column=col).coordinate}: se esperaba "
                f"'{formula_esperada}' y hay '{real}'"
            )
    # Porcentajes de indirectos/utilidad/IVA tomados de Variables
    for fila, esperado_f in ((ind, "=Variables!B12+Variables!B13"),
                             (util, "=Variables!B15"),
                             (iva, "=Variables!B18")):
        real = ws.cell(row=fila, column=6).value
        if real != esperado_f:
            errores.append(
                f"Presupuesto!{ws.cell(row=fila, column=6).coordinate}: se esperaba "
                f"'{esperado_f}' y hay '{real}'"
            )
    return errores


def _pu_oficiales() -> Optional[dict]:
    """PU oficial por clave, desde la BD canonica de costos.

    Devuelve None si la canonica no esta disponible: el check se omite y no
    se vuelve un error (los tests son autocontenidos).
    """
    try:
        from src.apu_db import _descubrir_db
        db = _descubrir_db()
    except Exception:
        db = None
    if not db or not Path(db).exists():
        return None
    try:
        import sqlite3
        con = sqlite3.connect(Path(db).as_uri() + "?mode=ro", uri=True)
        con.row_factory = sqlite3.Row
        filas = con.execute(
            "SELECT clave, precio_unitario, origen_precio FROM v_concepto_pu").fetchall()
        con.close()
    except Exception:
        return None
    return {r["clave"]: (r["precio_unitario"], r["origen_precio"]) for r in filas}


def _check_pu_oficial(wb, valores: Optional[dict]) -> list[str]:
    """Check 9: el PU cotizado debe coincidir con el oficial del catalogo.

    El modelo canonico define la prioridad (publicado > APU > manual). Si el
    libro cotiza otra cosa, se reporta como ADVERTENCIA con el impacto en
    dinero: cotizar el APU recalculado es legitimo, pero no puede quedar
    oculto.
    """
    oficiales = _pu_oficiales()
    if not oficiales:
        return []
    if "Presupuesto" not in wb.sheetnames:
        return []
    v = (valores or {}).get("Presupuesto", {})
    ws = wb["Presupuesto"]

    advertencias = []
    for r in _filas_concepto_presupuesto(ws):
        clave = ws.cell(row=r, column=2).value
        if not isinstance(clave, str) or not clave.strip():
            continue
        oficial = oficiales.get(clave.strip())
        if not oficial:
            continue
        pu_oficial, origen = oficial
        if pu_oficial is None:
            continue
        pu_oficial = float(pu_oficial)
        if pu_oficial <= 0:
            continue

        celda_pu = ws.cell(row=r, column=6).value
        if _es_formula(celda_pu):
            pu_cot = v.get(f"F{r}")
        else:
            pu_cot = celda_pu
        if not isinstance(pu_cot, (int, float)):
            continue
        pu_cot = float(pu_cot)
        if abs(pu_cot - pu_oficial) <= max(0.01, pu_oficial * 0.01):
            continue

        cantidad = ws.cell(row=r, column=4).value
        impacto = ""
        if isinstance(cantidad, (int, float)):
            impacto = f"; impacto ${(pu_cot - pu_oficial) * float(cantidad):,.2f}"
        dif_pct = (pu_cot - pu_oficial) / pu_oficial
        advertencias.append(
            f"Presupuesto!F{r}: el PU cotizado (${pu_cot:,.2f}) difiere del oficial "
            f"del catálogo (${pu_oficial:,.2f}, origen={origen}) en {dif_pct:+.1%}{impacto}"
        )
    return advertencias


def _check_totales_valores(wb, valores: Optional[dict]) -> tuple[list[str], list[str]]:
    """Check 7 (valores): cuadre aritmético con el libro recalculado."""
    if not valores:
        return [], ["No se pudo recalcular el libro (Excel no disponible): "
                    "el cuadre de totales no se verificó"]
    if "Presupuesto" not in valores:
        return [], ["Sin valores recalculados para la hoja Presupuesto"]

    errores = []
    ws = wb["Presupuesto"]
    pie = _pie_presupuesto(ws)
    faltan = [k for k in ("cd", "ind", "util", "sin_iva", "iva", "gran") if k not in pie]
    if faltan:
        return [], [f"No se verificaron totales (faltan filas: {', '.join(faltan)})"]

    v = valores.get("Presupuesto", {})

    def _num(fila, col="G"):
        val = v.get(f"{col}{fila}")
        return float(val) if isinstance(val, (int, float)) else None

    cd, ind, util, sin_iva, iva, gran = (pie["cd"], pie["ind"], pie["util"],
                                          pie["sin_iva"], pie["iva"], pie["gran"])
    vals = {k: _num(f) for k, f in (("cd", cd), ("ind", ind), ("util", util),
                                    ("sin_iva", sin_iva), ("iva", iva), ("gran", gran))}
    faltantes = [k for k, val in vals.items() if val is None]
    if faltantes:
        return [], [f"No se verificaron totales (sin valor recalculado: {', '.join(faltantes)})"]

    # Suma de importes de conceptos = costo directo
    suma_conceptos = 0.0
    for r in _filas_concepto_presupuesto(ws):
        imp = _num(r)
        if imp is not None:
            suma_conceptos += imp
    if abs(suma_conceptos - vals["cd"]) > 0.01:
        errores.append(
            f"El costo directo (${vals['cd']:,.2f}) no cuadra con la suma de los "
            f"importes de los conceptos (${suma_conceptos:,.2f})"
        )
    if abs((vals["cd"] + vals["ind"] + vals["util"]) - vals["sin_iva"]) > 0.01:
        errores.append("Total sin IVA no cuadra con CD + Indirectos + Utilidad")
    if abs((vals["sin_iva"] + vals["iva"]) - vals["gran"]) > 0.01:
        errores.append("Gran total no cuadra con Total sin IVA + IVA")

    # La Portada debe mostrar el mismo gran total
    if "Portada" in valores and "Portada" in wb.sheetnames:
        ws_p = wb["Portada"]
        v_port = valores.get("Portada", {})
        fila_total = None
        for r in range(1, ws_p.max_row + 1):
            if str(ws_p.cell(row=r, column=1).value).strip() == "Total":
                fila_total = r
                break
        if fila_total:
            v_total = v_port.get(f"F{fila_total}")
            if isinstance(v_total, (int, float)) and abs(float(v_total) - vals["gran"]) > 0.01:
                errores.append(
                    f"El Total de la Portada (${float(v_total):,.2f}) no coincide con "
                    f"el Gran Total del Presupuesto (${vals['gran']:,.2f})"
                )
        # La tabla de conceptos de la Portada debe sumar el costo directo
        suma_portada = 0.0
        filas_concepto = 0
        for r in range(1, ws_p.max_row + 1):
            etiqueta = ws_p.cell(row=r, column=1).value
            if not isinstance(etiqueta, str) or not etiqueta.strip():
                continue
            formula = ws_p.cell(row=r, column=6).value
            if not (isinstance(formula, str) and formula.startswith("=D")):
                continue
            valor = v_port.get(f"F{r}")
            if isinstance(valor, (int, float)):
                suma_portada += float(valor)
                filas_concepto += 1
        if filas_concepto and abs(suma_portada - vals["cd"]) > 0.01:
            errores.append(
                f"La tabla de la Portada suma ${suma_portada:,.2f} y el Costo Directo "
                f"del Presupuesto es ${vals['cd']:,.2f}"
            )
    # Ninguna cuadrilla puede quedar en $0: el IFERROR del VLOOKUP lo oculta y
    # el presupuesto queda subvaluado sin que se note.
    if "Insumos" in wb.sheetnames and "Insumos" in valores:
        ws_ins = wb["Insumos"]
        v_ins = valores.get("Insumos", {})
        for r in range(5, ws_ins.max_row + 1):
            tipo = str(ws_ins.cell(row=r, column=2).value or "").strip().lower()
            formula = ws_ins.cell(row=r, column=5).value
            if tipo != "cuadrilla" or not _es_formula(formula):
                continue
            valor = v_ins.get(f"E{r}")
            if isinstance(valor, (int, float)) and float(valor) <= 0:
                errores.append(
                    f"Insumos!E{r}: la cuadrilla '{ws_ins.cell(row=r, column=1).value}' "
                    f"resuelve $0 (no tiene $/Jornal en la hoja Cuadrillas): la mano "
                    f"de obra quedaría subvaluada"
                )
    return errores, []


# ---------------------------------------------------------------------------
# API principal
# ---------------------------------------------------------------------------

def validar_libro(ruta_xlsx: str, recalcular: bool = True) -> dict:
    """Valida la integridad del libro .xlsx.

    Args:
        ruta_xlsx: Ruta al libro generado.
        recalcular: Si True y hay Excel disponible, recalcula una copia para
                    verificar el cuadre de totales.

    Returns:
        {"archivo": str, "ok": bool, "errores": [str], "advertencias": [str],
         "checks": [{"nombre": str, "ok": bool, "detalle": str}]}
    """
    ruta = Path(ruta_xlsx)
    reporte = {
        "archivo": str(ruta),
        "ok": False,
        "errores": [],
        "advertencias": [],
        "checks": [],
    }

    if not ruta.exists():
        reporte["errores"].append(f"No existe el archivo: {ruta}")
        return reporte

    try:
        import openpyxl
        wb = openpyxl.load_workbook(ruta, data_only=False)
    except Exception as e:
        reporte["errores"].append(f"No se pudo abrir el libro: {e}")
        return reporte

    def _registrar(nombre, errores, nivel="error"):
        reporte["checks"].append({
            "nombre": nombre,
            "ok": not errores,
            "detalle": "; ".join(errores[:5]) + (" ..." if len(errores) > 5 else ""),
        })
        if errores:
            destino = reporte["errores"] if nivel == "error" else reporte["advertencias"]
            destino.extend(f"{nombre}: {e}" for e in errores)

    try:
        _registrar("1/9 Claves de VLOOKUP existentes en Insumos/Cuadrillas",
                   _check_vlookups(wb))
        _registrar("2/9 Claves de las hojas APU existentes en Insumos",
                   _check_apu_claves(wb))
        _registrar("2b/9 IDs locales de Insumos únicos",
                   _check_ids_locales(wb))
        _registrar("3/9 Categorías de Cuadrillas existentes en Calculo_categoria",
                   _check_categorias(wb))
        _registrar("4/9 Consistencia de unidades (Presupuesto / APU / Insumos)",
                   _check_unidades(wb))
        _registrar("5/9 Sin ciclos de referencia entre hojas",
                   _check_ciclos(wb))
        _registrar("6/9 Sintaxis válida de las fórmulas",
                   _check_formulas(wb))
        _registrar("7/9 Cadena de totales correctamente conectada",
                   _check_totales_estructura(wb))

        valores = _recalcular_excel(ruta) if recalcular else None
        errores_val, advertencias_val = _check_totales_valores(wb, valores)
        _registrar("8/9 Cuadre de totales (valores recalculados)",
                   errores_val)
        reporte["advertencias"].extend(advertencias_val)

        # 9: el PU cotizado debe coincidir con el oficial del catálogo.
        # Si la BD canónica no está disponible, el check se omite.
        _registrar("9/9 PU cotizado = PU oficial del catálogo",
                   _check_pu_oficial(wb, valores), nivel="advertencia")
    except Exception as e:  # el validador nunca debe tumbar la construcción
        reporte["errores"].append(f"Error inesperado en la validación: {e}")
    finally:
        try:
            wb.close()
        except Exception:
            pass

    reporte["ok"] = not reporte["errores"]
    return reporte


def formatear_reporte(reporte: dict) -> str:
    """Resumen legible del reporte de validación."""
    lineas = []
    lineas.append(f"Validación: {len(reporte['errores'])} errores, "
                  f"{len(reporte['advertencias'])} advertencias")
    for e in reporte["errores"]:
        lineas.append(f"  [ERROR] {e}")
    for a in reporte["advertencias"]:
        lineas.append(f"  [AVISO] {a}")
    return "\n".join(lineas)


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Uso: python -m src.validador <ruta.xlsx>")
        raise SystemExit(1)
    rep = validar_libro(sys.argv[1])
    print(formatear_reporte(rep))
    for c in rep["checks"]:
        print(f"  [{'OK ' if c['ok'] else 'ERR'}] {c['nombre']}"
              + (f" -> {c['detalle']}" if c["detalle"] else ""))
    raise SystemExit(0 if rep["ok"] else 1)
