"""
generar_plantilla.py
====================
Genera la plantilla base machote.xlsx con:
- Hojas FASAR con fórmulas nativas de Excel (no valores duros)
- Estructura de cuadrillas con fórmulas $/Jornal
- Hoja Variables con estructura de parámetros de proyecto
- Hojas de documentación (Notas)

La plantilla resultante es el insumo de presupuestar.py,
que la copia y llena con los datos del proyecto.

Uso:
    python generar_plantilla.py

Salida:
    plantillas/machote.xlsx
"""

import os
import sys
import datetime

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Agregar src al path para importar estilos
sys.path.insert(0, os.path.dirname(__file__))
from src.estilos import (
    AZUL_OSCURO, AZUL_HEADER, AZUL_MEDIO, GRIS_SECCION, BLANCO,
    NEGRO_TEXTO, ROJO_TEXTO, VERDE_CLARO, AMARILLO_CLARO,
    fill_azul_oscuro, fill_azul_header, fill_azul_medio, fill_gris_seccion,
    fill_blanco, fill_verde_claro, fill_amarillo_claro, fill_pu_venta,
    font_titulo, font_titulo_blanco, font_subtitulo,
    font_header_blanco, font_header_negro, font_header_col,
    font_seccion, font_normal, font_normal_rojo, font_negrita, font_italica,
    font_param_header, font_valor_editable, font_subtotal,
    align_center, align_left, align_right, align_center_nowrap,
    borde_fino, borde_medio_left,
    aplicar_borde, merge_y_escribir, escribir_celda, col_letra,
)

# =============================================================================
# CONFIGURACIÓN DE SALIDA
# =============================================================================

CARPETA_SALIDA = "plantillas"
ARCHIVO_SALIDA = os.path.join(CARPETA_SALIDA, "machote.xlsx")

# =============================================================================
# DATOS FIJOS DE LA PLANTILLA (no son del proyecto, son estructura)
# =============================================================================

CATEGORIAS_MO = [
    # (nombre, salario_nominal_diario)
    ("Oficial Albañil", 925),
    ("Ayudante General", 675),
    ("Oficial Fierrero", 1000),
    ("Oficial Carpintero", 1025),
    ("Oficial Electricista", 1075),
    ("Oficial Plomero", 1025),
    ("Soldador", 1175),
    ("Pailero", 1175),
    ("Cabo de oficios", 1250),
    ("Topógrafo", 1000),
    ("Ayudante Especialista", 750),
    ("Operador Maquinaria", 1000),
    ("Oficial Tubero", 1175),
    ("Oficial Instalador Especializado", 1025),
]

CUADRILLAS_BASE = [
    # (id, nombre, cant1, cat1, cant2, cat2, cant3, cat3, descripcion)
    ("CUAD001", "Cuadrilla tubería", 0.1, "Cabo de oficios", 1, "Oficial Tubero",
     1, "Ayudante Especialista", "Instalación de tubería"),
    ("CUAD002", "Cuadrilla instalación equipos", 0.1, "Cabo de oficios", 1,
     "Oficial Instalador Especializado", 1, "Ayudante Especialista", "Montaje de equipos"),
    ("CUAD003", "Cuadrilla albañilería básica", 0.1, "Cabo de oficios", 1,
     "Oficial Albañil", 1, "Ayudante General", "Trabajos de albañilería general"),
    ("CUAD004", "Cuadrilla fierrería", 0.1, "Cabo de oficios", 1,
     "Oficial Fierrero", 1, "Ayudante General", "Habilitado y armado de acero"),
    ("CUAD005", "Cuadrilla carpintería", 0.1, "Cabo de oficios", 1,
     "Oficial Carpintero", 1, "Ayudante General", "Cimbra y carpintería de obra"),
    ("CUAD006", "Cuadrilla soldadura", 0.1, "Cabo de oficios", 1,
     "Soldador", 1, "Ayudante Especialista", "Trabajos de soldadura"),
    ("CUAD007", "Cuadrilla topografía", 1, "Topógrafo", 1,
     "Ayudante General", None, None, "Levantamientos y trazo"),
    ("CUAD008", "Cuadrilla limpieza general", 0.1, "Cabo de oficios", 2,
     "Ayudante General", None, None, "Limpieza general de obra"),
    ("CUAD009", "Cuadrilla plomería", 0.1, "Cabo de oficios", 1,
     "Oficial Plomero", 1, "Ayudante Especialista", "Instalaciones hidrosanitarias"),
    ("CUAD010", "Cuadrilla electricista", 0.1, "Cabo de oficios", 1,
     "Oficial Electricista", 1, "Ayudante Especialista", "Instalaciones eléctricas"),
]

# Parámetros FASAR por defecto (editables en la plantilla)
FASAR_DEFAULTS = {
    "dias_calendario": 365,
    "descansos_semanales": 52,
    "descansos_obligatorios": 7,
    "vacaciones_dias": 12,
    "prima_vacacional": 0.25,
    "aguinaldo_dias": 15,
    "ausentismo_dias": 0,
    "uma_diaria": 117.31,
    "tope_sbc_uma": 25,
    "riesgo_trabajo": 0.0758875,
    "em_cuota_fija": 0.204,
    "em_excedente": 0.011,
    "em_prestaciones": 0.007,
    "em_gastos_medicos": 0.0105,
    "invalidez_vida": 0.0175,
    "guarderias": 0.01,
    "sar_retiro": 0.02,
    "cesantia_vejez": 0.0315,
    "infonavit": 0.05,
    "isn_estatal": 0.03,
}

# =============================================================================
# HOJA: PARAMETROS (FASAR)
# =============================================================================

FILA_PARAM = {}  # Guarda las filas de cada parámetro para referencia


def crear_hoja_parametros(wb):
    """Hoja Parametros con todos los parámetros FASAR editables.
    
    Las filas de cada parámetro se guardan en FILA_PARAM para que
    otras hojas puedan generar fórmulas que referencien estas celdas.
    """
    ws = wb.create_sheet("Parametros")
    r = 1  # contador de fila

    # ---- Título ----
    merge_y_escribir(ws, f"A{r}:C{r}",
                     "PARÁMETROS FASAR — MÉXICO 2026",
                     font=font_titulo_blanco, fill=fill_azul_oscuro, alignment=align_center)
    for c in range(1, 4):
        ws.cell(row=r, column=c).fill = fill_azul_oscuro
    r += 1

    merge_y_escribir(ws, f"A{r}:C{r}",
                     "Edita las celdas en azul. Las fórmulas en Categorias y Calculo_categoria se actualizan automáticamente.",
                     font=font_italica, alignment=align_left)
    r += 2

    # ---- Sección 1: Días de trabajo ----
    merge_y_escribir(ws, f"A{r}:C{r}", "1. DÍAS DE TRABAJO",
                     font=font_param_header, fill=fill_azul_header, alignment=align_left)
    aplicar_borde(ws, r, 1, 3)
    r += 1

    params_dias = [
        ("dias_calendario", "Días calendario del año", 365),
        ("descansos_semanales", "Descansos semanales pagados", 52),
        ("descansos_obligatorios", "Días de descanso obligatorio pagados", 7),
        ("vacaciones_dias", "Vacaciones pagadas — 1er año (días)", 12),
        ("prima_vacacional", "Prima vacacional mínima (%)", 0.25),
        ("aguinaldo_dias", "Aguinaldo mínimo (días)", 15),
        ("ausentismo_dias", "Ausentismo / improductividad estimada (días)", 0),
    ]

    for key, label, valor in params_dias:
        escribir_celda(ws, r, 1, label, font=font_normal, alignment=align_left)
        escribir_celda(ws, r, 2, valor, font=font_valor_editable, alignment=align_center)
        FILA_PARAM[key] = r
        r += 1

    r += 1  # espacio

    # ---- Pre-asignar filas para fórmulas derivadas ----
    claves_derivadas = [
        "Tl_dias_laborales",
        "Tp_dias_pagados",
        "Te_dias_fuera_base",
        "Tp_menos_Te",
        "factor_Tp_Tl",
        "factor_integracion",
    ]
    for key in claves_derivadas:
        FILA_PARAM[key] = r
        r += 1

    # ---- Ahora escribir las fórmulas derivadas (ya todas las filas existen) ----
    fc = FILA_PARAM  # alias para legibilidad

    formulas_dias = [
        ("Tl_dias_laborales", "Tl — Días efectivamente laborados",
         f"=B{fc['dias_calendario']}-B{fc['descansos_semanales']}-B{fc['descansos_obligatorios']}-B{fc['vacaciones_dias']}-B{fc['ausentismo_dias']}"),
        ("Tp_dias_pagados", "Tp — Días realmente pagados",
         f"=B{fc['dias_calendario']}+B{fc['aguinaldo_dias']}+B{fc['vacaciones_dias']}*B{fc['prima_vacacional']}"),
        ("Te_dias_fuera_base", "Te — Días pagados fuera de la base de cargas",
         f"=B{fc['vacaciones_dias']}"),
        ("Tp_menos_Te", "Tp − Te (días para cargas)",
         f"=B{fc['Tp_dias_pagados']}-B{fc['Te_dias_fuera_base']}"),
        ("factor_Tp_Tl", "Factor Tp / Tl",
         f"=B{fc['Tp_dias_pagados']}/B{fc['Tl_dias_laborales']}"),
        ("factor_integracion", "Factor de integración LFT (FI = SDI / Sn)",
         f"=(B{fc['dias_calendario']}+B{fc['aguinaldo_dias']}+B{fc['vacaciones_dias']}*B{fc['prima_vacacional']})/B{fc['dias_calendario']}"),
    ]

    for key, label, formula in formulas_dias:
        fila = FILA_PARAM[key]
        escribir_celda(ws, fila, 1, label, font=font_negrita, alignment=align_left)
        escribir_celda(ws, fila, 2, formula, font=font_negrita, alignment=align_center,
                       number_format="0.0000")

    r += 2

    # ---- Sección 2: UMA y topes ----
    merge_y_escribir(ws, f"A{r}:C{r}", "2. UMA Y TOPES",
                     font=font_param_header, fill=fill_azul_header, alignment=align_left)
    aplicar_borde(ws, r, 1, 3)
    r += 1

    uma_params = [
        ("uma_diaria", "UMA diaria vigente ($)", 117.31),
        ("tope_sbc_uma", "Tope SBC (veces UMA)", 25),
    ]
    for key, label, valor in uma_params:
        escribir_celda(ws, r, 1, label, font=font_normal, alignment=align_left)
        escribir_celda(ws, r, 2, valor, font=font_valor_editable, alignment=align_center,
                       number_format="#,##0.00")
        FILA_PARAM[key] = r
        r += 1

    # SBC tope derivado
    escribir_celda(ws, r, 1, "SBC tope ($) = UMA × Tope",
                   font=font_negrita, alignment=align_left)
    escribir_celda(ws, r, 2,
                   f"=B{FILA_PARAM['uma_diaria']}*B{FILA_PARAM['tope_sbc_uma']}",
                   font=font_negrita, alignment=align_center, number_format="#,##0.00")
    FILA_PARAM["sbc_tope"] = r
    r += 2

    # ---- Sección 3: Cuotas patronales ----
    merge_y_escribir(ws, f"A{r}:C{r}", "3. CUOTAS PATRONALES (tasas editables)",
                     font=font_param_header, fill=fill_azul_header, alignment=align_left)
    aplicar_borde(ws, r, 1, 3)
    r += 1

    cuotas = [
        ("riesgo_trabajo", "Riesgo de trabajo (% del SBC)", 0.0758875),
        ("em_cuota_fija", "E&M cuota fija (% de 1 UMA diaria)", 0.204),
        ("em_excedente", "E&M excedente de 3 UMA (% del excedente)", 0.011),
        ("em_prestaciones", "E&M prestaciones en dinero (% del SBC)", 0.007),
        ("em_gastos_medicos", "E&M gastos médicos pensionados (% del SBC)", 0.0105),
        ("invalidez_vida", "Invalidez y vida (% del SBC)", 0.0175),
        ("guarderias", "Guarderías y prestaciones sociales (% del SBC)", 0.01),
        ("sar_retiro", "SAR / Retiro (% del SBC)", 0.02),
        ("cesantia_vejez", "Cesantía y vejez (% del SBC)", 0.0315),
        ("infonavit", "Infonavit (% del SBC)", 0.05),
        ("isn_estatal", "ISN estatal (% del SBC)", 0.03),
    ]

    for key, label, valor in cuotas:
        escribir_celda(ws, r, 1, label, font=font_normal, alignment=align_left)
        escribir_celda(ws, r, 2, valor, font=font_valor_editable, alignment=align_center,
                       number_format="0.0000")
        FILA_PARAM[key] = r
        r += 1

    r += 1

    # Suma de % SBC (sin E&M fija ni excedente, que se calculan aparte)
    # Las cuotas que son % del SBC: riesgo, prestaciones, gastos médicos, invalidez,
    # guarderías, SAR, cesantía, infonavit, ISN
    filas_pct_sbc = [
        FILA_PARAM["riesgo_trabajo"],
        FILA_PARAM["em_prestaciones"],
        FILA_PARAM["em_gastos_medicos"],
        FILA_PARAM["invalidez_vida"],
        FILA_PARAM["guarderias"],
        FILA_PARAM["sar_retiro"],
        FILA_PARAM["cesantia_vejez"],
        FILA_PARAM["infonavit"],
        FILA_PARAM["isn_estatal"],
    ]
    suma_formula = "=" + "+".join(f"B{f}" for f in filas_pct_sbc)

    escribir_celda(ws, r, 1, "SUMA % SBC (cuotas sobre nómina)",
                   font=font_negrita, alignment=align_left)
    escribir_celda(ws, r, 2, suma_formula,
                   font=font_negrita, alignment=align_center, number_format="0.0000")
    FILA_PARAM["suma_pct_sbc"] = r

    # ---- Anchos de columna ----
    ws.column_dimensions["A"].width = 48
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 20

    # ---- Congelar panel ----
    ws.freeze_panes = "A2"

    return ws


# =============================================================================
# HOJA: CATEGORIAS
# =============================================================================


def crear_hoja_categorias(wb):
    """Hoja Categorias con las categorías de mano de obra y su Sn editable."""
    ws = wb.create_sheet("Categorias")

    # Título
    merge_y_escribir(ws, "A1:D1",
                     "CATEGORÍAS DE MANO DE OBRA — Salario nominal diario (Sn)",
                     font=font_titulo_blanco, fill=fill_azul_oscuro, alignment=align_center)
    for c in range(1, 5):
        ws.cell(row=1, column=c).fill = fill_azul_oscuro

    merge_y_escribir(ws, "A2:D2",
                     "Edita el salario nominal (Sn) de cada categoría. El FASAR y Sr se recalculan en Calculo_categoria.",
                     font=font_italica, alignment=align_left)
    r = 4

    # Headers
    headers = ["Categoría", "Sn — Salario nominal diario ($)", "Salario semanal ($)", "ID categoría"]
    for i, h in enumerate(headers, 1):
        escribir_celda(ws, r, i, h, font=font_header_blanco, fill=fill_azul_medio,
                       alignment=align_center)
    aplicar_borde(ws, r, 1, 4)
    r += 1

    # Datos
    for idx, (nombre, sn) in enumerate(CATEGORIAS_MO, 1):
        cat_id = f"CAT{idx:03d}"
        escribir_celda(ws, r, 1, nombre, font=font_normal, alignment=align_left)
        escribir_celda(ws, r, 2, sn, font=font_normal_rojo, alignment=align_center,
                       number_format="#,##0.00")
        # Salario semanal = Sn × 7 (aproximado, para referencia)
        escribir_celda(ws, r, 3, f"=B{r}*7", font=font_normal, alignment=align_center,
                       number_format="#,##0.00")
        escribir_celda(ws, r, 4, cat_id, font=font_normal, alignment=align_center)
        aplicar_borde(ws, r, 1, 4)
        r += 1

    # Anchos
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 14

    ws.freeze_panes = "A5"
    return ws


# =============================================================================
# HOJA: CALCULO_CATEGORIA (FASAR por categoría)
# =============================================================================


def crear_hoja_calculo_categoria(wb):
    """Hoja Calculo_categoria con el cálculo FASAR por categoría usando fórmulas.

    Referencias:
    - Parametros: UMA, tope SBC, cuotas, Tp/Tl, FI
    - Categorias: Sn de cada categoría

    Columnas:
    A: Categoría
    B: Sn (salario nominal)
    C: SDI = Sn × FI
    D: SBC = min(SDI, tope_SBC)
    E: E&M cuota fija ($)
    F: E&M excedente 3 UMA ($)
    G: Cargas % SBC ($)
    H: Total cargas diarias ($)
    I: Ps (obligaciones / SBC)
    J: FASAR = (Tp/Tl) × (1 + Ps)
    K: Sr = Sn × FASAR ($/jornal)
    """
    ws = wb.create_sheet("Calculo_categoria")

    # ---- Parámetros de referencia ----
    f_tptl = FILA_PARAM["factor_Tp_Tl"]      # fila en Parametros
    f_fi = FILA_PARAM["factor_integracion"]
    f_uma = FILA_PARAM["uma_diaria"]
    f_tope = FILA_PARAM["sbc_tope"]
    f_em_fija = FILA_PARAM["em_cuota_fija"]
    f_em_exc = FILA_PARAM["em_excedente"]
    f_suma = FILA_PARAM["suma_pct_sbc"]

    # ---- Título ----
    merge_y_escribir(ws, "A1:K1",
                     "SALARIOS Y FASAR POR CATEGORÍA — Sueldo base editable + cálculo automático",
                     font=font_titulo_blanco, fill=fill_azul_oscuro, alignment=align_center)
    for c in range(1, 12):
        ws.cell(row=1, column=c).fill = fill_azul_oscuro

    merge_y_escribir(ws, "A2:K2",
                     "Edita el Sueldo base (col B). El FASAR y Sr se calculan solos con los parámetros FASAR.",
                     font=font_italica, alignment=align_left)
    r = 4

    # ---- Headers ----
    headers_calc = [
        ("Categoría", 32),
        ("Sn ($)", 12),
        ("SDI ($)", 12),
        ("SBC con tope ($)", 14),
        ("E&M cuota fija ($)", 14),
        ("E&M excedente ($)", 14),
        ("Cargas % SBC ($)", 14),
        ("Total cargas ($)", 14),
        ("Ps", 10),
        ("FASAR", 10),
        ("Sr — Salario real ($/jornal)", 18),
    ]
    for i, (hdr, ancho) in enumerate(headers_calc, 1):
        escribir_celda(ws, r, i, hdr, font=font_header_blanco, fill=fill_azul_medio,
                       alignment=align_center)
        ws.column_dimensions[col_letra(i)].width = ancho
    aplicar_borde(ws, r, 1, 11)
    r += 1

    # ---- Filas de datos (una por categoría) ----
    for idx in range(len(CATEGORIAS_MO)):
        nombre, sn = CATEGORIAS_MO[idx]

        # A: Categoría
        escribir_celda(ws, r, 1, nombre, font=font_normal, alignment=align_left)

        # B: Sueldo base (EDITABLE, directo aquí)
        escribir_celda(ws, r, 2, sn, font=font_normal_rojo,
                       alignment=align_center, number_format="#,##0.00")

        # C: SDI = Sn × FI
        escribir_celda(ws, r, 3, f"=B{r}*Parametros!B{f_fi}",
                       font=font_normal, alignment=align_center, number_format="#,##0.00")

        # D: SBC = min(SDI, tope_UMA)
        escribir_celda(ws, r, 4, f"=MIN(C{r},Parametros!B{f_tope})",
                       font=font_normal, alignment=align_center, number_format="#,##0.00")

        # E: E&M cuota fija $ = tasa_fija × UMA
        escribir_celda(ws, r, 5, f"=Parametros!B{f_em_fija}*Parametros!B{f_uma}",
                       font=font_normal, alignment=align_center, number_format="#,##0.00")

        # F: E&M excedente $ = if(SBC > 3×UMA, (SBC − 3×UMA) × tasa_excedente, 0)
        escribir_celda(ws, r, 6,
                       f"=IF(D{r}>3*Parametros!B{f_uma},(D{r}-3*Parametros!B{f_uma})*Parametros!B{f_em_exc},0)",
                       font=font_normal, alignment=align_center, number_format="#,##0.00")

        # G: Cargas % SBC $ = SBC × suma_tasas
        escribir_celda(ws, r, 7, f"=D{r}*Parametros!B{f_suma}",
                       font=font_normal, alignment=align_center, number_format="#,##0.00")

        # H: Total cargas $ = E + F + G
        escribir_celda(ws, r, 8, f"=E{r}+F{r}+G{r}",
                       font=font_normal, alignment=align_center, number_format="#,##0.00")

        # I: Ps = Total cargas / SBC
        escribir_celda(ws, r, 9, f"=IF(D{r}>0,H{r}/D{r},0)",
                       font=font_normal, alignment=align_center, number_format="0.0000")

        # J: FASAR = (Tp/Tl) × (1 + Ps)
        escribir_celda(ws, r, 10, f"=Parametros!B{f_tptl}*(1+I{r})",
                       font=font_negrita, alignment=align_center, number_format="0.0000",
                       fill=fill_verde_claro)

        # K: Sr = Sn × FASAR
        escribir_celda(ws, r, 11, f"=B{r}*J{r}",
                       font=font_negrita, alignment=align_center, number_format="#,##0.00",
                       fill=fill_verde_claro)

        aplicar_borde(ws, r, 1, 11)
        r += 1

    # ---- Nota al pie ----
    r += 1
    merge_y_escribir(ws, f"A{r}:K{r}",
                     "Para agregar una categoría: inserta una fila en Categorias y copia cualquier fila de esta hoja. "
                     "Las fórmulas se ajustan automáticamente.",
                     font=font_italica, alignment=align_left)

    ws.freeze_panes = "A5"
    return ws


# =============================================================================
# HOJA: CUADRILLAS
# =============================================================================


def crear_hoja_cuadrillas(wb):
    """Hoja Cuadrillas — FUENTE MAESTRA de costos por jornal.
    
    Aquí se definen las cuadrillas y su $/Jornal EDITABLE directamente.
    Insumos hará VLOOKUP a esta hoja para obtener el $/Jornal de cada cuadrilla.
    APUs referencian a Insumos, que a su vez referencian a Cuadrillas.
    
    Estructura:
    A: ID Cuadrilla | B: Nombre | C: $/Jornal (EDITABLE) | D: Descripción | E: Cant1 | F: Cat1 | G: Cant2 | H: Cat2 | I: Cant3 | J: Cat3 | K: Composición (texto)
    """
    ws = wb.create_sheet("Cuadrillas")

    # ---- Título ----
    merge_y_escribir(ws, "A1:K1",
                     "CUADRILLAS — Costos por jornal (FUENTE MAESTRA EDITABLE)",
                     font=DNA().font_titulo(), fill=DNA().fill_azul_oscuro(), alignment=DNA().align_center())
    for c in range(1, 12):
        ws.cell(row=1, column=c).fill = DNA().fill_azul_oscuro()
    ws.row_dimensions[1].height = 30

    merge_y_escribir(ws, "A2:K2",
                     "FUENTE MAESTRA: Edite directamente el $/Jornal (col C). Insumos y APUs referencian esta hoja.",
                     font=DNA().font_cuerpo(bold=True), alignment=align_left)
    r = 3

    # ---- Headers ----
    headers_cuad = [
        ("ID Cuadrilla", 14),
        ("Nombre Cuadrilla", 30),
        ("$/Jornal", 14),      # ← EDITABLE DIRECTAMENTE
        ("Descripción", 40),
        ("Cant. 1", 10),
        ("Categoría 1", 28),
        ("Cant. 2", 10),
        ("Categoría 2", 28),
        ("Cant. 3", 10),
        ("Categoría 3", 28),
        ("Composición (texto)", 50),
    ]
    for i, (hdr, ancho) in enumerate(headers_cuad, 1):
        escribir_celda(ws, 3, i, hdr, font=DNA().font_header(), 
                       fill=DNA().fill_header_azul(), alignment=DNA().align_center())
        ws.column_dimensions[get_column_letter(i)].width = ancho
    aplicar_borde(ws, 3, 1, 11)
    r = 4

    # ---- Filas de cuadrillas predefinidas (con $/Jornal calculado como valor inicial) ----
    for cuad in CUADRILLAS_BASE:
        cuad_id, nombre, c1, cat1, c2, cat2, c3, cat3, desc = cuad

        # Calcular $/Jornal inicial desde Calculo_categoria (valor inicial)
        # Pero la celda C será EDITABLE por el usuario
        partes = [
            f"C{r}*VLOOKUP(D{r},Calculo_categoria!A5:K50,11,FALSE)",
            f"E{r}*VLOOKUP(F{r},Calculo_categoria!A5:K50,11,FALSE)",
        ]
        if c3 is not None and cat3 is not None:
            partes.append(f"G{r}*VLOOKUP(H{r},Calculo_categoria!A5:K50,11,FALSE)")

        formula_jornal = "=" + "+".join(partes)

        escribir_celda(ws, r, 1, cuad_id, font=fuente_normal, alignment=align_center)
        escribir_celda(ws, r, 2, nombre, font=fuente_normal, alignment=align_left)
        # Col C = $/Jornal - VALOR INICIAL (formula) PERO EDITABLE
        escribir_celda(ws, r, 3, formula_jornal, 
                       font=fuente_normal, alignment=align_right, number_format="$#,##0.00")
        escribir_celda(ws, r, 4, desc, font=fuente_normal, alignment=align_left)
        escribir_celda(ws, r, 5, c1, font=fuente_normal, alignment=align_center, number_format="0.00")
        escribir_celda(ws, r, 6, cat1, font=fuente_normal, alignment=align_left)
        escribir_celda(ws, r, 7, c2, font=fuente_normal, alignment=align_center, number_format="0.00")
        escribir_celda(ws, r, 8, cat2, font=fuente_normal, alignment=align_left)
        if c3 is not None:
            escribir_celda(ws, r, 10, c3, font=fuente_normal, alignment=align_center, number_format="0.00")
        if cat3 is not None:
            escribir_celda(ws, r, 11, cat3, font=fuente_normal, alignment=align_left)
        
        # Composición texto
        comp = f"{c1}x{cat1}"
        if c2: comp += f" + {c2}x{cat2}"
        if c3 and cat3: comp += f" + {c3}x{cat3}"
        escribir_celda(ws, r, 12, comp, font=fuente_italica, alignment=align_left)

        aplicar_borde(ws, r, 1, 12)
        r += 1

    # Fila de instrucción
    r += 1
    merge_y_escribir(ws, f"A{r}:L{r}",
                     "INSTRUCCIÓN: Modifique directamente la columna C ($/Jornal) para actualizar costos. "
                     "Insumos y APUs se recalculan automáticamente.",
                     font=fuente_italica, alignment=align_left)

        aplicar_borde(ws, r, 1, 10)
        r += 1

    # ---- Nota ----
    r += 1
    merge_y_escribir(ws, f"A{r}:J{r}",
                     "El rango de VLOOKUP (A5:K50) cubre hasta 46 categorías. "
                     "Si agregas más categorías en Calculo_categoria, amplía el rango en la fórmula.",
                     font=font_italica, alignment=align_left)

    ws.freeze_panes = "A5"
    return ws


# =============================================================================
# HOJA: VARIABLES (parámetros del proyecto)
# =============================================================================


def crear_hoja_variables(wb):
    """Hoja Variables: parámetros globales del proyecto, editables.

    Los valores por defecto son placeholders. presupuestar.py los
    sobrescribe con los datos del proyecto real.
    """
    ws = wb.create_sheet("Variables")

    # ---- Título ----
    merge_y_escribir(ws, "A1:C1",
                     "VARIABLES DEL PROYECTO — Editables",
                     font=font_titulo_blanco, fill=fill_azul_oscuro, alignment=align_center)
    for c in range(1, 4):
        ws.cell(row=1, column=c).fill = fill_azul_oscuro

    merge_y_escribir(ws, "A2:C2",
                     "Modifica los valores de esta hoja. Todas las hojas APU y Presupuesto los referencian.",
                     font=font_italica, alignment=align_left)
    r = 4

    # ---- Headers ----
    for i, h in enumerate(["Variable", "Valor", "Notas"], 1):
        escribir_celda(ws, r, i, h, font=font_header_blanco, fill=fill_azul_medio,
                       alignment=align_center)
    aplicar_borde(ws, r, 1, 3)
    r += 1

    # ---- Datos del proyecto ----
    merge_y_escribir(ws, f"A{r}:C{r}", "DATOS DEL PROYECTO",
                     font=font_param_header, fill=fill_azul_header, alignment=align_left)
    aplicar_borde(ws, r, 1, 3)
    r += 1

    proyecto_vars = [
        ("Nombre del proyecto", "PLACEHOLDER — Aeropuerto Internacional", "Texto libre"),
        ("Clave del proyecto", "PLACEHOLDER — PB-OMA-R01", "Texto libre"),
        ("Cliente / Empresa", "PLACEHOLDER — OMA S.A. de C.V.", "Texto libre"),
        ("Fecha de presupuesto", datetime.date.today().strftime("%Y-%m-%d"), "YYYY-MM-DD"),
    ]
    for label, valor, nota in proyecto_vars:
        escribir_celda(ws, r, 1, label, font=font_normal, alignment=align_left)
        escribir_celda(ws, r, 2, valor, font=font_valor_editable, alignment=align_left)
        escribir_celda(ws, r, 3, nota, font=font_italica, alignment=align_left)
        aplicar_borde(ws, r, 1, 3)
        r += 1

    r += 1

    # ---- Indirectos y utilidad ----
    merge_y_escribir(ws, f"A{r}:C{r}", "INDIRECTOS Y UTILIDAD",
                     font=font_param_header, fill=fill_azul_header, alignment=align_left)
    aplicar_borde(ws, r, 1, 3)
    r += 1

    financ_vars = [
        ("% Indirectos de campo (CIC)", 0.14, "Fracción del Costo Directo"),
        ("% Indirectos de oficina (CIO)", 0.05, "Fracción del Costo Directo"),
        ("% Financiamiento", 0.0135, "Fracción de (CD + CI)"),
        ("% Utilidad", 0.08, "Fracción de (CD + CI + Fin)"),
    ]
    for label, valor, nota in financ_vars:
        escribir_celda(ws, r, 1, label, font=font_normal, alignment=align_left)
        escribir_celda(ws, r, 2, valor, font=font_valor_editable, alignment=align_center,
                       number_format="0.00%")
        escribir_celda(ws, r, 3, nota, font=font_italica, alignment=align_left)
        aplicar_borde(ws, r, 1, 3)
        r += 1

    r += 1

    # ---- Impuestos y otros ----
    merge_y_escribir(ws, f"A{r}:C{r}", "IMPUESTOS Y OTROS",
                     font=font_param_header, fill=fill_azul_header, alignment=align_left)
    aplicar_borde(ws, r, 1, 3)
    r += 1

    impuesto_vars = [
        ("% IVA", 0.16, "Impuesto al valor agregado"),
        ("% ISN estatal", 0.03, "Impuesto sobre nómina (referencia, ya está en FASAR)"),
        ("Rendimiento default", 1.0, "Factor de rendimiento para mano de obra"),
    ]
    for label, valor, nota in impuesto_vars:
        escribir_celda(ws, r, 1, label, font=font_normal, alignment=align_left)
        escribir_celda(ws, r, 2, valor, font=font_valor_editable, alignment=align_center,
                       number_format="0.00%")
        escribir_celda(ws, r, 3, nota, font=font_italica, alignment=align_left)
        aplicar_borde(ws, r, 1, 3)
        r += 1

    # ---- Anchos ----
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 42

    ws.freeze_panes = "A5"
    return ws


# =============================================================================
# HOJA: NOTAS
# =============================================================================


def crear_hoja_notas(wb):
    """Hoja Notas con alcance, fundamento legal y referencias."""
    ws = wb.create_sheet("Notas")

    merge_y_escribir(ws, "A1:B1",
                     "ALCANCE Y FUNDAMENTO — Presupuestador PONTEM",
                     font=font_titulo_blanco, fill=fill_azul_oscuro, alignment=align_center)

    notas = [
        ("Objetivo",
         "Plantilla paramétrica para presupuestos de obra en México. "
         "Calcula el costo de mano de obra considerando todas las obligaciones patronales "
         "conforme a la LFT, Ley del IMSS e Infonavit."),
        ("FASAR / FSR (art. 191 RLOPSRM)",
         "Fsr = (Tp / Tl) × (1 + Ps). Te se toma como los días de vacaciones pagadas. "
         "Ps son las obligaciones obrero-patronales como fracción del SBC de cada categoría."),
        ("Salario real (art. 190 RLOPSRM)",
         "Sr = Sn × Fsr. El costo directo de mano de obra debe considerar el salario real "
         "integrado, que incluye todas las prestaciones y obligaciones patronales."),
        ("Tp (días pagados)",
         "Días realmente pagados en el año: 365 + aguinaldo + prima vacacional."),
        ("Tl (días laborados)",
         "Días efectivamente laborados: días calendario − descansos semanales − "
         "descansos obligatorios − vacaciones − ausentismo."),
        ("Prestaciones mínimas LFT (1er año)",
         "12 días de vacaciones, prima vacacional mínima 25%, 15 días de aguinaldo."),
        ("SDI legal",
         "Factor de integración LFT/SDI = (365 + aguinaldo + vacaciones × prima_vacacional) ÷ 365."),
        ("Ps variable por categoría",
         "No existe un Ps único 'por ley'. Varía por prima de riesgo, ISN estatal, "
         "y SBC de cada trabajador. Esta plantilla calcula un Ps específico para cada categoría."),
        ("Exclusiones",
         "No incluye: herramienta menor, EPP, transporte, viáticos, tiempo extra, "
         "prima dominical, ni prestaciones contractuales adicionales. "
         "Estos se agregan como insumos en cada APU."),
        ("Explosión de insumos (BOM)",
         "La hoja Explosion (generada por presupuestar.py) consolida todos los insumos "
         "de todos los APUs: cantidad_total = Σ (cantidad_obra × consumo_unitario). "
         "Sirve para compras, programación de suministros y contratación de mano de obra."),
        ("Revisión profesional",
         "Valida los porcentajes de cargas patronales vigentes con tu contador "
         "o especialista en seguridad social antes de presentar una licitación."),
    ]

    r = 3
    for titulo, contenido in notas:
        escribir_celda(ws, r, 1, titulo, font=font_negrita, alignment=align_left)
        escribir_celda(ws, r, 2, contenido, font=font_normal,
                       alignment=Alignment(horizontal="left", vertical="center", wrap_text=True))
        r += 1

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 100

    return ws


# =============================================================================
# FUNCIÓN PRINCIPAL
# =============================================================================


def generar_plantilla():
    """Genera la plantilla base machote.xlsx."""
    print(f"\n{'='*60}")
    print(f"  GENERANDO PLANTILLA BASE (machote.xlsx)")
    print(f"  Salida: {ARCHIVO_SALIDA}")
    print(f"{'='*60}\n")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # 1. Parametros (FASAR) — hoja oculta de parámetros
    print("  [1/4] Creando Parametros (FASAR)...")
    crear_hoja_parametros(wb)

    # 2. Calculo_categoria (salario base editable + FASAR) — visible
    print("  [2/4] Creando Calculo_categoria (salario + FASAR)...")
    crear_hoja_calculo_categoria(wb)

    # 3. Cuadrillas (depende de Calculo_categoria)
    print("  [3/4] Creando Cuadrillas...")
    crear_hoja_cuadrillas(wb)

    # 4. Variables (parámetros del proyecto)
    print("  [4/4] Creando Variables...")
    crear_hoja_variables(wb)

    # Guardar
    os.makedirs(CARPETA_SALIDA, exist_ok=True)
    print(f"\n  Guardando: {ARCHIVO_SALIDA} ...")
    wb.save(ARCHIVO_SALIDA)

    size_kb = os.path.getsize(ARCHIVO_SALIDA) / 1024
    print(f"\n{'='*60}")
    print(f"  PLANTILLA GENERADA")
    print(f"  Archivo: {os.path.abspath(ARCHIVO_SALIDA)}")
    print(f"  Tamaño:  {size_kb:.1f} KB")
    print(f"  Hojas:   {wb.sheetnames}")
    print(f"{'='*60}\n")

    return ARCHIVO_SALIDA


if __name__ == "__main__":
    generar_plantilla()