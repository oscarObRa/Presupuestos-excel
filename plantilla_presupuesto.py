"""
plantilla_presupuesto.py
=========================
Genera una plantilla Excel unificada de presupuesto de obra, combinando:

1. Formato/colores → Catálogo muro rev 2.xlsx (APU sheets)
2. FASAR/Calculadora → Calculadora_FASAR_Mexico_2026_corregida.xlsx
3. Estructura APU + insumos → OMA-CUU-VIP-PCI-XXX-CAT-XXX-0 cotizacion rev1.xlsx

Uso:
    python plantilla_presupuesto.py

El archivo de salida se guarda como "Plantilla_Presupuesto_OMA.xlsx" en el mismo directorio.
Las variables editables están en la sección CONFIGURACION al inicio del script.
"""

import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, NamedStyle, numbers
)
from openpyxl.utils import get_column_letter
from copy import copy
import datetime
import os

# =============================================================================
# CONFIGURACION EDITABLE
# =============================================================================

# Datos del proyecto
PROYECTO = "AEROPUERTO INTERNACIONAL - SALA VIP OMA"
CLAVE_PROYECTO = "PB - OMA - R01"
EMPRESA = "OMA / AEROPUERTO INTERNACIONAL, S.A. DE C.V."
FECHA = datetime.date.today().strftime("%d de %B de %Y").upper()

# Variables de precio (editables)
INDIRECTO_CAMPO = 0.14       # CIC = % del CD
INDIRECTO_OFICINA = 0.05     # CIO = % del CD
FINANCIAMIENTO = 0.0135      # F = % del (CD + CI)
UTILIDAD = 0.08              # U = % del (CD + CI + F)

# Conceptos/APUs a generar (cada uno será una hoja)
CONCEPTOS_APU = [
    {"codigo": "APU-01", "descripcion": "DESCRIPCIÓN DEL CONCEPTO 01", "unidad": "PZA"},
    {"codigo": "APU-02", "descripcion": "DESCRIPCIÓN DEL CONCEPTO 02", "unidad": "ML"},
    {"codigo": "APU-03", "descripcion": "DESCRIPCIÓN DEL CONCEPTO 03", "unidad": "M2"},
    {"codigo": "APU-04", "descripcion": "DESCRIPCIÓN DEL CONCEPTO 04", "unidad": "M3"},
    {"codigo": "APU-05", "descripcion": "DESCRIPCIÓN DEL CONCEPTO 05", "unidad": "KG"},
    {"codigo": "APU-06", "descripcion": "DESCRIPCIÓN DEL CONCEPTO 06", "unidad": "LOTE"},
    {"codigo": "APU-07", "descripcion": "DESCRIPCIÓN DEL CONCEPTO 07", "unidad": "ML"},
    {"codigo": "APU-08", "descripcion": "DESCRIPCIÓN DEL CONCEPTO 08", "unidad": "PZA"},
    {"codigo": "APU-09", "descripcion": "DESCRIPCIÓN DEL CONCEPTO 09", "unidad": "SISTEMA"},
    {"codigo": "APU-10", "descripcion": "DESCRIPCIÓN DEL CONCEPTO 10", "unidad": "LOTE"},
]

# Filas vacías mínimas por sección en cada APU
FILAS_VACIAS_POR_SECCION = 6

# Archivo de salida
CARPETA_SALIDA = "plantillas"
ARCHIVO_SALIDA = os.path.join(CARPETA_SALIDA, "Plantilla_Presupuesto_OMA.xlsx")

# =============================================================================
# COLORES Y ESTILOS (basados en Catálogo muro rev 2)
# =============================================================================

# Colores principales
AZUL_OSCURO = "1F3864"       # Encabezados principales
AZUL_HEADER = "DDEBF7"       # Headers de columna en APU (tipo OMA)
GRIS_SECCION = "F2F2F2"     # Fondo de títulos de sección
BLANCO = "FFFFFF"
NEGRO_TEXTO = "000000"
ROJO_TEXTO = "FF0000"

# Fills
fill_azul_oscuro = PatternFill(start_color=AZUL_OSCURO, end_color=AZUL_OSCURO, fill_type="solid")
fill_azul_header = PatternFill(start_color=AZUL_HEADER, end_color=AZUL_HEADER, fill_type="solid")
fill_gris_seccion = PatternFill(start_color=GRIS_SECCION, end_color=GRIS_SECCION, fill_type="solid")
fill_blanco = PatternFill(start_color=BLANCO, end_color=BLANCO, fill_type="solid")

# Fonts
font_titulo_principal = Font(name="Calibri", size=11, bold=True, color=NEGRO_TEXTO)
font_header_azul_oscuro = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
font_header_azul_claro = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
font_seccion_bold = Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO)
font_normal = Font(name="Calibri", size=10, color=NEGRO_TEXTO)
font_normal_rojo = Font(name="Calibri", size=11, color=ROJO_TEXTO)
font_header_col = Font(name="Calibri", size=12, bold=True, color=NEGRO_TEXTO)
font_subtotal = Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO)
font_pu_venta = Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO)
font_label_bold = Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO)

# Alignments
align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_right = Alignment(horizontal="right", vertical="center")
align_center_center = Alignment(horizontal="center", vertical="center")

# Borders
borde_fino = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)
borde_medio_left = Border(left=Side(style="medium"))
borde_inferior = Border(bottom=Side(style="thin"))
borde_superior_grueso = Border(top=Side(style="thick"))
borde_doble_inferior = Border(bottom=Side(style="double"))

# =============================================================================
# FUNCIONES AUXILIARES
# =============================================================================

def aplicar_borde(ws, fila, col_inicio, col_fin, estilo="thin"):
    """Aplica bordes finos a un rango horizontal."""
    for col in range(col_inicio, col_fin + 1):
        cell = ws.cell(row=fila, column=col)
        cell.border = Border(
            left=Side(style=estilo), right=Side(style=estilo),
            top=Side(style=estilo), bottom=Side(style=estilo)
        )


def merge_and_set(ws, rango, valor, font=None, fill=None, alignment=None):
    """Hace merge de un rango y asigna valor + estilos."""
    ws.merge_cells(rango)
    cell = ws[rango.split(":")[0]]
    cell.value = valor
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment


def set_cell(ws, fila, col, valor, font=None, fill=None, alignment=None, number_format=None):
    """Asigna valor y estilos a una celda."""
    cell = ws.cell(row=fila, column=col)
    cell.value = valor
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment
    if number_format:
        cell.number_format = number_format


# =============================================================================
# HOJAS DE LA CALCULADORA FASAR
# =============================================================================

def crear_hoja_categorias(wb):
    """Crea la hoja Categorias (salarios base por categoría)."""
    ws = wb.create_sheet("Categorias")

    categorias = [
        ("Oficial Albañil", 925, 5500),
        ("Ayudante General", 675, 4000),
        ("Oficial Fierrero", 1000, 6000),
        ("Oficial Carpintero", 1025, 6200),
        ("Oficial Electricista", 1075, 6500),
        ("Oficial Plomero", 1025, 6200),
        ("Soldador", 1175, 7000),
        ("Pailero", 1175, 7000),
        ("Cabo de oficios", 1250, 7500),
        ("Topógrafo", 1000, 6000),
        ("Ayudante Especialista", 750, 4500),
        ("Operador Maquinaria", 1000, 6000),
        ("Oficial Tubero", 1175, 7000),
        ("Oficial Instalador Especializado", 1025, 6200),
    ]

    # Headers
    headers = ["Categoría", "Salario nominal diario (Sn)", "Salario semanal"]
    for i, h in enumerate(headers, 1):
        set_cell(ws, 1, i, h, font=font_header_azul_oscuro,
                 fill=fill_azul_header, alignment=align_center)

    for idx, (cat, sn, ss) in enumerate(categorias, 2):
        set_cell(ws, idx, 1, cat, font=font_normal, alignment=align_left)
        set_cell(ws, idx, 2, sn, font=font_normal_rojo, alignment=align_center)
        set_cell(ws, idx, 3, ss, font=font_normal_rojo, alignment=align_center)

    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 18

    return ws


def crear_hoja_parametros(wb):
    """Crea la hoja Parametros con los parámetros del FASAR."""
    ws = wb.create_sheet("Parametros")

    params = [
        ("PARÁMETROS GENERALES — EDITABLES", None, True),
        ("", None, False),
        ("1. Relación de Días al Año", None, True),
        ("Año de análisis", 2026, False),
        ("Días calendario del año", 365, False),
        ("Descansos semanales pagados", 52, False),
        ("Días de descanso obligatorio pagados", 7, False),
        ("Vacaciones pagadas (1er año)", 12, False),
        ("Prima vacacional mínima", 0.25, False),
        ("Aguinaldo mínimo (días)", 15, False),
        ("Otros días pagados contractuales", 0, False),
        ("Ausentismo / improductividad estimada (días)", 0, False),
        ("Días laborales teóricos", 294, False),
        ("Días efectivamente laborados (Tl)", 294, False),
        ("", None, False),
        ("2. Factor de Prestaciones (Tp / Tl)", None, True),
        ("Días realmente pagados (Tp)", 383, False),
        ("Días pagados para cargas (Tp − Te)", 371, False),
        ("Ps manual (solo se usa si Modo de Ps = Manual)", 0.30, False),
        ("FASAR con Ps manual (solo referencia)", 1.6813, False),
        ("Factor de integración LFT / SDI", 1.0493, False),
        ("SBC diario — categoría seleccionada (con tope)", None, False),
        ("", None, False),
        ("3. Cuotas patronales — tasas editables (se aplican al SBC de cada categoría)", None, True),
        ("UMA diaria vigente", 117.31, False),
        ("Prima de riesgo de trabajo", 0.0758875, False),
        ("IMSS – Riesgo de trabajo (% del SBC)", 0.0758875, False),
        ("IMSS – E&M cuota fija (% de 1 UMA diaria)", 0.204, False),
        ("IMSS – E&M excedente de 3 UMA (% del excedente)", 0.011, False),
        ("IMSS – E&M prestaciones en dinero (% del SBC)", 0.007, False),
        ("IMSS – E&M gastos médicos pensionados (% del SBC)", 0.0105, False),
        ("IMSS – Invalidez y vida (% del SBC)", 0.0175, False),
        ("IMSS – Guarderías y prestaciones sociales (% del SBC)", 0.01, False),
        ("SAR / Retiro (% del SBC)", 0.02, False),
        ("Cesantía y vejez patronal", "Tabla progresiva (ver abajo)", False),
        ("Infonavit (% del SBC)", 0.05, False),
        ("ISN estatal (% del SBC)", 0.03, False),
        ("Total cargas patronales diarias — categoría seleccionada", None, False),
        ("Ps calculado — categoría seleccionada", None, False),
        ("", None, False),
        ("Modo de Ps (Calculado / Manual)", "Calculado", False),
        ("Te: días pagados fuera de la base de cargas", 12, False),
        ("Tope de SBC (veces UMA)", 25, False),
    ]

    for i, (label, valor, is_header) in enumerate(params, 1):
        font = Font(name="Calibri", size=10, bold=is_header, color=NEGRO_TEXTO)
        set_cell(ws, i, 1, label, font=font, alignment=align_left)
        if valor is not None:
            set_cell(ws, i, 2, valor, font=Font(name="Calibri", size=10, color=NEGRO_TEXTO),
                     alignment=align_center)

    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 30

    return ws


def crear_hoja_calculadora(wb):
    """Crea la hoja Calculadora (FASAR interactivo)."""
    ws = wb.create_sheet("Calculadora")

    merge_and_set(ws, "A1:E1", "CALCULADORA DE SALARIO REAL (FASAR) — MÉXICO 2026",
                  font=Font(name="Calibri", size=14, bold=True, color="FFFFFF"),
                  fill=fill_azul_oscuro, alignment=align_center)

    merge_and_set(ws, "A2:E2",
                  "Selecciona la categoría; el sueldo, las cuotas patronales y el FASAR se calculan automáticamente.",
                  font=Font(name="Calibri", size=9, italic=True, color="404040"),
                  alignment=align_left)

    datos = [
        ("Categoría", "A", True),
        ("Salario nominal diario (Sn)", "B", True),
        ("Factor de salario real (FASAR / FSR)", "C", True),
        ("Salario real para APU (Sr = Sn × FASAR)", "D", True),
        ("Factor de integración LFT / SDI", "E", True),
        ("Salario diario integrado legal (SDI = Sn × FI)", "F", True),
        ("Salario base de cotización (SBC, con tope)", "G", True),
    ]

    for i, (label, col, is_param) in enumerate(datos, 4):
        set_cell(ws, i, 1, label,
                 font=Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO),
                 fill=fill_azul_header, alignment=align_left)
        set_cell(ws, i, 2, None, font=font_normal, alignment=align_center,
                 fill=PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"))

    # Parámetros extra (derecha)
    extra_datos = [
        ("Días efectivamente laborados (Tl)", 294),
        ("Días realmente pagados (Tp)", 383),
        ("Ps: obligaciones patronales", None),
        ("Prima vacacional", 0.25),
        ("Aguinaldo (días)", 15),
        ("Vacaciones (días)", 12),
        ("Cargas patronales diarias ($)", None),
    ]
    for i, (label, valor) in enumerate(extra_datos, 4):
        set_cell(ws, i, 4, label,
                 font=Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO),
                 fill=fill_azul_header, alignment=align_left)
        if valor is not None:
            set_cell(ws, i, 5, valor, font=font_normal, alignment=align_center,
                     fill=PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"))

    # Fórmula
    merge_and_set(ws, "A12:E12", "FÓRMULA EMPLEADA POR EL FASAR",
                  font=Font(name="Calibri", size=11, bold=True, color="FFFFFF"),
                  fill=fill_azul_header, alignment=align_center)

    merge_and_set(ws, "A13:E13",
                  'FASAR = (Tp / Tl) × [1 + Ps × ((Tp − Te) / Tp)]     con Te = vacaciones (editable en hoja Parametros)',
                  font=Font(name="Calibri", size=10, color=NEGRO_TEXTO),
                  alignment=align_left)

    merge_and_set(ws, "A15:E15",
                  "Interpretación: Sr es el costo de mano de obra por jornada realmente trabajada para el contratista.",
                  font=Font(name="Calibri", size=9, italic=True, color="404040"),
                  alignment=align_left)

    merge_and_set(ws, "A17:E17",
                  'Para cambiar salarios, edita la hoja "Categorias". Para adecuar cargas, calendario e impuestos, edita la hoja "Parametros".',
                  font=Font(name="Calibri", size=9, italic=True, color="404040"),
                  alignment=align_left)

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 34
    ws.column_dimensions["E"].width = 16

    return ws


def crear_hoja_cuadrillas(wb):
    """Crea la hoja Cuadrillas con las cuadrillas predefinidas."""
    ws = wb.create_sheet("Cuadrillas")

    # Título
    merge_and_set(ws, "B2:J2", "CUADRILLAS — costos por jornal con SALARIO REAL (FASAR)",
                  font=Font(name="Calibri", size=12, bold=True, color=NEGRO_TEXTO),
                  alignment=align_left)

    # Headers
    headers = ["Id", "Cuadrilla", "Cantidad 1", "Categoría 1", "Cantidad 2",
               "Categoría 2", "Cantidad 3", "Categoría 3", "Costo/Jornal", "Descripción"]
    for i, h in enumerate(headers, 1):
        set_cell(ws, 4, i, h,
                 font=Font(name="Calibri", size=10, bold=True, color="FFFFFF"),
                 fill=fill_azul_header, alignment=align_center)

    cuadrillas = [
        ("CUAD001", "Cuadrilla tubería", 0.1, "Cabo de oficios", 1, "Oficial Tubero",
         1, "Ayudante Especialista"),
        ("CUAD002", "Cuadrilla instalación equipos", 0.1, "Cabo de oficios", 1,
         "Oficial Instalador Especializado", 1, "Ayudante Especialista"),
        ("CUAD003", "Cuadrilla albañilería básica", 0.1, "Cabo de oficios", 1,
         "Oficial Albañil", 1, "Ayudante General"),
        ("CUAD004", "Cuadrilla fierrería", 0.1, "Cabo de oficios", 1,
         "Oficial Fierrero", 1, "Ayudante General"),
        ("CUAD005", "Cuadrilla carpintería", 0.1, "Cabo de oficios", 1,
         "Oficial Carpintero", 1, "Ayudante General"),
        ("CUAD006", "Cuadrilla soldadura", 0.1, "Cabo de oficios", 1,
         "Soldador", 1, "Ayudante Especialista"),
        ("CUAD007", "Cuadrilla topografía", 1, "Topógrafo", 1,
         "Ayudante General", None, None),
        ("CUAD008", "Cuadrilla limpieza general", 0.1, "Cabo de oficios", 2,
         "Ayudante General", None, None),
        ("CUAD009", "Cuadrilla plomería", 0.1, "Cabo de oficios", 1,
         "Oficial Plomero", 1, "Ayudante Especialista"),
        ("CUAD010", "Cuadrilla electricista", 0.1, "Cabo de oficios", 1,
         "Oficial Electricista", 1, "Ayudante Especialista"),
    ]

    for idx, c in enumerate(cuadrillas, 5):
        set_cell(ws, idx, 1, c[0], font=font_normal, alignment=align_center)
        set_cell(ws, idx, 2, c[1], font=font_normal, alignment=align_left)
        set_cell(ws, idx, 3, c[2], font=font_normal, alignment=align_center)
        set_cell(ws, idx, 4, c[3], font=font_normal, alignment=align_left)
        set_cell(ws, idx, 5, c[4], font=font_normal, alignment=align_center)
        set_cell(ws, idx, 6, c[5], font=font_normal, alignment=align_left)
        if c[6] is not None:
            set_cell(ws, idx, 7, c[6], font=font_normal, alignment=align_center)
        if c[7] is not None:
            set_cell(ws, idx, 8, c[7], font=font_normal, alignment=align_left)

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 32
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 32
    ws.column_dimensions["G"].width = 12
    ws.column_dimensions["H"].width = 28
    ws.column_dimensions["I"].width = 18
    ws.column_dimensions["J"].width = 60

    return ws


def crear_hoja_notas(wb):
    """Crea la hoja Notas con alcance y fundamento."""
    ws = wb.create_sheet("Notas")

    notas = [
        ("ALCANCE Y FUNDAMENTO", "", True, True),
        ("Objetivo", "Plantilla paramétrica para presupuestos de obra en México, anualizada a 365 días. "
         "Calcula el costo de mano de obra considerando todas las obligaciones patronales conforme a la LFT, "
         "Ley del IMSS e Infonavit vigentes para 2026.", True, False),
        ("Salario real (Sr)",
         "Sr = Sn × Fsr. El artículo 190 del RLOPSRM indica que el costo directo de mano de obra "
         "debe considerar el salario real integrado, que incluye todas las prestaciones y obligaciones patronales.",
         True, False),
        ("FASAR / FSR",
         "Fsr = (Tp / Tl) × [1 + Ps × ((Tp − Te) / Tp)]. En esta plantilla Te se toma como los días de "
         "vacaciones pagadas (editables en hoja Parametros).",
         True, False),
        ("Tp",
         "Días realmente pagados en el año: días calendario + aguinaldo + prima vacacional.",
         True, False),
        ("Tl",
         "Días efectivamente laborados: días calendario menos descansos semanales, días de descanso "
         "obligatorio, vacaciones y ausentismo estimado.",
         True, False),
        ("Prestaciones mínimas LFT",
         "Para 1 año de servicio: 12 días de vacaciones; prima vacacional mínima de 25%; y 15 días de aguinaldo.",
         True, False),
        ("SDI legal",
         "El factor de integración LFT/SDI se calcula como 1 + aguinaldo/365 + (vacaciones × prima vacacional)/365.",
         True, False),
        ("Ps no es fijo",
         "No existe un Ps único 'por ley' aplicable a todas las empresas. Varía por prima de riesgo, "
         "ISN estatal, y salario de cada trabajador. Esta calculadora obtiene un Ps específico para cada categoría.",
         True, False),
        ("Exclusiones",
         "No incluye herramienta menor, EPP, transporte, viáticos, tiempo extra, prima dominical, "
         "ni otras prestaciones contractuales. Estos se agregan como insumos adicionales en cada APU.",
         True, False),
        ("Revisión profesional",
         "Antes de presentar una licitación o determinar cuotas, valida los porcentajes de cargas "
         "patronales vigentes con tu contador o especialista en seguridad social.",
         True, False),
        ("Ps por categoría",
         'Con Modo de Ps = Calculado (Parametros!B38), el Ps se calcula para cada categoría usando las '
         'cuotas patronales acumuladas y el SBC correspondiente.',
         True, False),
    ]

    for i, (titulo, contenido, is_title, _) in enumerate(notas, 1):
        font_t = Font(name="Calibri", size=11, bold=True, color=NEGRO_TEXTO)
        set_cell(ws, i, 1, titulo, font=font_t, alignment=align_left)
        if contenido:
            set_cell(ws, i, 2, contenido, font=font_normal, alignment=Alignment(horizontal="left", vertical="center", wrap_text=True))

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 110

    return ws


def crear_hoja_calculo_categoria(wb):
    """Crea la hoja Calculo_categoria con el FASAR por categoría."""
    ws = wb.create_sheet("Calculo_categoria")

    merge_and_set(ws, "B2:S2",
                  "FASAR POR CATEGORÍA — se alimenta de la hoja Categorias (columnas B a E son el resultado; F a S son el desglose de cuotas)",
                  font=Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO),
                  alignment=align_left)

    # Headers
    headers_calc = ["Categoria", "Sueldo", "FASAR", "Sr (Sueldo real)", "SBC (con tope)",
                    "E&M cuota fija", "E&M excedente 3 UMA", "E&M prest. dinero"]
    for i, h in enumerate(headers_calc, 2):  # Start at col B
        set_cell(ws, 4, i, h,
                 font=Font(name="Calibri", size=9, bold=True, color="FFFFFF"),
                 fill=PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid"),
                 alignment=align_center)

    categorias_fasar = [
        ("Oficial Albañil", 925, 1.7162),
        ("Ayudante General", 675, 1.7259),
        ("Oficial Fierrero", 1000, 1.7143),
        ("Oficial Carpintero", 1025, 1.7137),
        ("Oficial Electricista", 1075, 1.7126),
        ("Oficial Plomero", 1025, 1.7137),
        ("Soldador", 1175, 1.7107),
        ("Pailero", 1175, 1.7107),
        ("Cabo de oficios", 1250, 1.7094),
        ("Topógrafo", 1000, 1.7143),
        ("Ayudante Especialista", 750, 1.7223),
        ("Operador Maquinaria", 1000, 1.7143),
        ("Oficial Tubero", 1175, 1.7107),
        ("Oficial Instalador Especializado", 1025, 1.7137),
    ]

    for idx, (cat, sueldo, fasar) in enumerate(categorias_fasar, 5):
        set_cell(ws, idx, 2, cat, font=font_normal, alignment=align_left)
        set_cell(ws, idx, 3, sueldo, font=font_normal, alignment=align_center)
        set_cell(ws, idx, 4, fasar, font=font_normal, alignment=align_center,
                 fill=PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid"))
        # Sr = sueldo * fasar
        set_cell(ws, idx, 5, round(sueldo * fasar, 2), font=font_normal, alignment=align_center,
                 fill=PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid"))

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 18

    return ws


# =============================================================================
# HOJA DE RESUMEN (insumos)
# =============================================================================

def crear_hoja_insumos(wb, codigos_apu):
    """Crea la hoja insumos que consolida todos los insumos de las hojas APU."""
    ws = wb.create_sheet("insumos")

    # Headers
    headers = ["Tipo", "ID", "Nombre insumo", "Cantidad", "PU", "Total", "APU Origen"]
    for i, h in enumerate(headers, 1):
        set_cell(ws, 1, i, h,
                 font=Font(name="Calibri", size=10, bold=True, color="FFFFFF"),
                 fill=fill_azul_header, alignment=align_center)

    # Filas de ejemplo por tipo (Material, Mano de Obra, Equipo, Auxiliar)
    ejemplos = [
        ("Material", "MAT-01", "Descripción del material 1", None, None, None, codigos_apu[0] if codigos_apu else ""),
        ("Material", "MAT-02", "Descripción del material 2", None, None, None, ""),
        ("Material", "MAT-03", "Descripción del material 3", None, None, None, ""),
        ("Mano de Obra", "CUAD001", "Cuadrilla de trabajo 1", None, None, None, ""),
        ("Mano de Obra", "CUAD002", "Cuadrilla de trabajo 2", None, None, None, ""),
        ("Equipo", "EQP-01", "Equipo / herramienta 1", None, None, None, ""),
        ("Equipo", "EQP-02", "Equipo / herramienta 2", None, None, None, ""),
        ("Auxiliar", "AUX-01", "Servicio auxiliar 1", None, None, None, ""),
        ("Auxiliar", "AUX-02", "Servicio auxiliar 2", None, None, None, ""),
    ]

    for idx, (tipo, ins_id, nombre, cant, pu, total, origen) in enumerate(ejemplos, 2):
        set_cell(ws, idx, 1, tipo, font=font_normal, alignment=align_center)
        set_cell(ws, idx, 2, ins_id, font=font_normal, alignment=align_center)
        set_cell(ws, idx, 3, nombre, font=font_normal, alignment=align_left)
        set_cell(ws, idx, 4, cant, font=font_normal, alignment=align_center,
                 number_format="#,##0.00")
        set_cell(ws, idx, 5, pu, font=font_normal, alignment=align_right,
                 number_format="$#,##0.00")
        set_cell(ws, idx, 6, total, font=font_normal, alignment=align_right,
                 number_format="$#,##0.00")
        set_cell(ws, idx, 7, origen, font=font_normal, alignment=align_center)

    # Nota al pie
    nota_row = len(ejemplos) + 3
    merge_and_set(ws, f"A{nota_row}:G{nota_row}",
                  "▶ Esta hoja se llena automáticamente al copiar los insumos de cada APU. "
                  "Cada fila referencia el insumo, su cantidad total, PU y la hoja APU de origen.",
                  font=Font(name="Calibri", size=9, italic=True, color="404040"),
                  alignment=align_left)

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 55
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 14

    return ws


# =============================================================================
# HOJA DE VARIABLES DE PRECIO
# =============================================================================

def crear_hoja_variables(wb):
    """Crea la hoja 'variable' con los porcentajes editables (CIC, CIO, Fin, Util)."""
    ws = wb.create_sheet("variable")

    merge_and_set(ws, "A1:C1",
                  "VARIABLES DE PRECIO (editables - afectan a todas las hojas de PU)",
                  font=Font(name="Calibri", size=11, bold=True, color=NEGRO_TEXTO),
                  alignment=align_left)

    # Headers
    for i, h in enumerate(["Concepto", "Valor", "Aplicación"], 1):
        set_cell(ws, 3, i, h,
                 font=Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO),
                 fill=fill_azul_header, alignment=align_center)

    variables = [
        ("Indirecto de campo (CIC)", INDIRECTO_CAMPO, "% sobre Costo Directo"),
        ("Indirecto de oficina central (CIO)", INDIRECTO_OFICINA, "% sobre Costo Directo"),
        ("Financiamiento (F)", FINANCIAMIENTO, "% sobre (CD + CI)"),
        ("Utilidad (U)", UTILIDAD, "% sobre (CD + CI + Financiamiento)"),
    ]

    for idx, (concepto, valor, aplicacion) in enumerate(variables, 4):
        set_cell(ws, idx, 1, concepto, font=font_normal, alignment=align_left)
        set_cell(ws, idx, 2, valor, font=Font(name="Calibri", size=11, bold=True, color=NEGRO_TEXTO),
                 alignment=align_center, number_format="0.00%")
        set_cell(ws, idx, 3, aplicacion, font=font_normal, alignment=align_left)

    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 40

    return ws


# =============================================================================
# HOJA DE PRESUPUESTO RESUMEN
# =============================================================================

def crear_hoja_presupuesto(wb, conceptos_apu):
    """Crea la hoja 'Presupuesto' con el resumen del catálogo de conceptos."""
    ws = wb.create_sheet("Presupuesto")

    # Título
    merge_and_set(ws, "A1:F1",
                  f"CATÁLOGO DE CONCEPTOS — {PROYECTO}",
                  font=Font(name="Calibri", size=14, bold=True, color=NEGRO_TEXTO),
                  alignment=align_center)
    merge_and_set(ws, "A2:F2", EMPRESA,
                  font=Font(name="Calibri", size=11, bold=True, color=NEGRO_TEXTO),
                  alignment=align_center)
    merge_and_set(ws, "A4:F4", f"Fecha de Actualización: {FECHA}",
                  font=Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO),
                  alignment=align_left)

    # Headers
    headers = ["Part.", "Concepto", "Unidad", "Cantidad", "P.U.", "Importe"]
    for i, h in enumerate(headers, 1):
        set_cell(ws, 7, i, h,
                 font=Font(name="Calibri", size=10, bold=True, color="FFFFFF"),
                 fill=PatternFill(start_color="003366", end_color="003366", fill_type="solid"),
                 alignment=align_center)

    # Filas de conceptos
    for idx, apu in enumerate(conceptos_apu, 8):
        set_cell(ws, idx, 1, apu["codigo"], font=font_normal, alignment=align_center)
        set_cell(ws, idx, 2, apu["descripcion"], font=font_normal, alignment=align_left)
        set_cell(ws, idx, 3, apu["unidad"], font=font_normal, alignment=align_center)
        set_cell(ws, idx, 4, None, font=font_normal, alignment=align_center, number_format="#,##0.00")
        set_cell(ws, idx, 5, None, font=font_normal, alignment=align_right, number_format="$#,##0.00")
        set_cell(ws, idx, 6, None, font=Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO),
                 alignment=align_right, number_format="$#,##0.00")
        aplicar_borde(ws, idx, 1, 6, "thin")

    # Total
    total_row = 8 + len(conceptos_apu)
    merge_and_set(ws, f"A{total_row}:E{total_row}", "SUMA TOTAL",
                  font=Font(name="Calibri", size=11, bold=True, color="FFFFFF"),
                  fill=PatternFill(start_color="000000", end_color="000000", fill_type="solid"),
                  alignment=align_center)
    set_cell(ws, total_row, 6, None,
             font=Font(name="Calibri", size=11, bold=True, color="FFFFFF"),
             fill=PatternFill(start_color="000000", end_color="000000", fill_type="solid"),
             alignment=align_right, number_format="$#,##0.00")

    # Gran total con IVA
    total_iva_row = total_row + 1
    merge_and_set(ws, f"A{total_iva_row}:E{total_iva_row}",
                  "GRAN TOTAL (IMPUESTO INCLUIDO: 16%)",
                  font=Font(name="Calibri", size=11, bold=True, color="FFFFFF"),
                  fill=PatternFill(start_color="000000", end_color="000000", fill_type="solid"),
                  alignment=align_center)
    set_cell(ws, total_iva_row, 6, None,
             font=Font(name="Calibri", size=11, bold=True, color="FFFFFF"),
             fill=PatternFill(start_color="000000", end_color="000000", fill_type="solid"),
             alignment=align_right, number_format="$#,##0.00")

    # Notas
    nota_row = total_iva_row + 3
    set_cell(ws, nota_row, 1, "Notas:", font=Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO))
    set_cell(ws, nota_row + 1, 1, " - Vigencia de esta cotización: 30 días",
             font=font_normal, alignment=align_left)
    set_cell(ws, nota_row + 2, 1, " - Precios en MXN (pesos mexicanos), no incluyen IVA salvo indicación expresa.",
             font=font_normal, alignment=align_left)
    set_cell(ws, nota_row + 3, 1, " - Modelo de contrato a precio alzado. Los volúmenes son de referencia.",
             font=font_normal, alignment=align_left)

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 18

    return ws


# =============================================================================
# HOJAS APU (Análisis de Precio Unitario)
# =============================================================================

def crear_hoja_apu(wb, codigo, descripcion, unidad, filas_vacias=6):
    """
    Crea una hoja APU con el formato del Catálogo muro rev 2.
    Cada sección (Materiales, Mano de Obra, Equipo, Maquinaria, Auxiliares)
    tiene al menos `filas_vacias` renglones vacíos.
    """
    ws = wb.create_sheet(codigo)

    # ---- ROW 1: Encabezado ----
    # Primero el merge de A1:B1 y ponemos la empresa en A1 (celda principal del merge)
    ws.merge_cells("A1:B1")
    set_cell(ws, 1, 1, EMPRESA, font=font_titulo_principal, alignment=align_left)
    set_cell(ws, 1, 3, codigo,
             font=Font(name="Calibri", size=10, bold=True, color="FFFFFF"),
             fill=fill_azul_oscuro, alignment=align_center)
    merge_and_set(ws, f"D1:F1", f"Concepto:\n{descripcion}",
                  font=Font(name="Calibri", size=9, bold=True, color="FFFFFF"),
                  fill=fill_azul_oscuro, alignment=Alignment(horizontal="left", vertical="center", wrap_text=True))

    # ---- ROW 2: Espacio ----
    ws.row_dimensions[2].height = 8

    # ---- ROW 3: Datos del concepto ----
    set_cell(ws, 3, 2, PROYECTO, font=font_normal, alignment=align_left)
    set_cell(ws, 3, 3, "Unidad:", font=font_label_bold, alignment=Alignment(horizontal="right", vertical="center"))
    set_cell(ws, 3, 4, unidad, font=font_normal, alignment=align_right)
    set_cell(ws, 3, 5, "PU:", font=font_label_bold, alignment=Alignment(horizontal="right", vertical="center"))
    # PU fórmula: =F43 (el PU DE VENTA)
    set_cell(ws, 3, 6, None, font=font_subtotal, alignment=align_right, number_format="$#,##0.00")

    # ---- ROW 4: Clave del proyecto y cantidad ----
    set_cell(ws, 4, 2, CLAVE_PROYECTO, font=font_normal, alignment=Alignment(horizontal="left", vertical="center"))
    set_cell(ws, 4, 3, "Cantidad:", font=font_label_bold, alignment=Alignment(horizontal="right", vertical="center"))
    set_cell(ws, 4, 4, None, font=font_normal, alignment=align_right, number_format="#,##0.00")
    set_cell(ws, 4, 5, "Importe:", font=font_label_bold, alignment=Alignment(horizontal="right", vertical="center"))
    set_cell(ws, 4, 6, None, font=font_subtotal, alignment=align_right, number_format="$#,##0.00")

    # ---- ROW 5: Headers de tabla ----
    headers_apu = ["ID", "Insumo", "Unidad", "Cantidad", "P. unitario", "Importe"]
    for i, h in enumerate(headers_apu, 1):
        set_cell(ws, 5, i, h,
                 font=font_header_col, alignment=align_center)

    # ---- SECCIONES ----
    # Layout:
    #   Fila de título de sección
    #   + filas_vacias filas para datos
    #
    # Secciones:
    #   1. Materiales
    #   2. Mano de Obra
    #   3. Equipo
    #   4. Maquinaria
    #   5. Auxiliares

    secciones = [
        ("Materiales", "MAT"),
        ("Mano de Obra", "MO"),
        ("Equipo", "EQP"),
        ("Maquinaria", "MAQ"),
        ("Auxiliares", "AUX"),
    ]

    fila_actual = 6  # Empezamos después de los headers
    seccion_info = {}  # Guardar: {nombre: (fila_inicio, fila_fin)}

    for sec_num, (nombre, prefijo) in enumerate(secciones, 1):
        # Fila de título de sección
        seccion_info[nombre] = {"inicio": fila_actual, "fin": fila_actual + filas_vacias}

        set_cell(ws, fila_actual, 1, sec_num, font=font_seccion_bold, alignment=align_center)
        merge_and_set(ws, f"B{fila_actual}:E{fila_actual}", nombre,
                      font=font_seccion_bold, fill=fill_gris_seccion,
                      alignment=align_left)
        # Total de sección (con fórmula SUMA)
        total_formula = f"=SUM(F{fila_actual + 1}:F{fila_actual + filas_vacias})"
        set_cell(ws, fila_actual, 6, total_formula,
                 font=font_seccion_bold, fill=fill_gris_seccion,
                 alignment=align_right, number_format="$#,##0.00")

        aplicar_borde(ws, fila_actual, 1, 6, "thin")

        # Filas vacías para datos
        for r_offset in range(1, filas_vacias + 1):
            r = fila_actual + r_offset
            for c in range(1, 7):
                cell = ws.cell(row=r, column=c)
                cell.font = font_normal
                cell.alignment = align_center if c in (1, 3, 4) else align_left if c == 2 else align_right
                if c in (5, 6):
                    cell.number_format = "$#,##0.00"
                elif c == 4:
                    cell.number_format = "#,##0.0000"
                cell.border = Border(
                    left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"), bottom=Side(style="thin")
                )
            # Borde izquierdo medio (estilo Catálogo muro)
            ws.cell(row=r, column=1).border = Border(
                left=Side(style="medium"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin")
            )

        fila_actual += filas_vacias + 1  # +1 por la fila de título

    # ---- SUBTOTALES Y TOTALES ----
    # COSTO DIRECTO (CD) = SUMA de todos los totales de sección
    cd_formula_parts = []
    for nombre, info in seccion_info.items():
        cd_formula_parts.append(f"F{info['inicio']}")
    cd_formula = "=" + "+".join(cd_formula_parts)

    # CD: merge C:E, valor en F
    merge_and_set(ws, f"C{fila_actual}:E{fila_actual}", "COSTO DIRECTO (CD)",
                  font=font_subtotal, fill=fill_gris_seccion, alignment=align_left)
    set_cell(ws, fila_actual, 6, cd_formula,
             font=font_subtotal, fill=fill_gris_seccion,
             alignment=align_right, number_format="$#,##0.00")
    aplicar_borde(ws, fila_actual, 3, 6, "thin")
    cd_row = fila_actual
    fila_actual += 1

    # CIC: merge C:D, valor % en E, resultado en F
    merge_and_set(ws, f"C{fila_actual}:D{fila_actual}", "Indirectos Campo (CIC = %CD)",
                  font=font_normal, alignment=align_left)
    set_cell(ws, fila_actual, 5, INDIRECTO_CAMPO, font=Font(name="Calibri", size=11, color=ROJO_TEXTO),
             alignment=align_center, number_format="0%")
    set_cell(ws, fila_actual, 6, f"=F{cd_row}*E{fila_actual}",
             font=font_normal, alignment=align_right, number_format="$#,##0.00")
    aplicar_borde(ws, fila_actual, 3, 6, "thin")
    cic_row = fila_actual
    fila_actual += 1

    # CIO: merge C:D, valor % en E, resultado en F
    merge_and_set(ws, f"C{fila_actual}:D{fila_actual}", "Indirectos Oficina Central (CIO = %CD)",
                  font=font_normal, alignment=align_left)
    set_cell(ws, fila_actual, 5, INDIRECTO_OFICINA, font=Font(name="Calibri", size=11, color=ROJO_TEXTO),
             alignment=align_center, number_format="0%")
    set_cell(ws, fila_actual, 6, f"=F{cd_row}*E{fila_actual}",
             font=font_normal, alignment=align_right, number_format="$#,##0.00")
    aplicar_borde(ws, fila_actual, 3, 6, "thin")
    cio_row = fila_actual
    fila_actual += 1

    # CI: merge C:E, formula en F
    merge_and_set(ws, f"C{fila_actual}:E{fila_actual}", "TOTAL INDIRECTOS (CI = CIC + CIO)",
                  font=font_subtotal, fill=fill_gris_seccion, alignment=align_left)
    set_cell(ws, fila_actual, 6, f"=F{cic_row}+F{cio_row}",
             font=font_subtotal, fill=fill_gris_seccion,
             alignment=align_right, number_format="$#,##0.00")
    aplicar_borde(ws, fila_actual, 3, 6, "thin")
    ci_row = fila_actual
    fila_actual += 1

    # Espacio
    fila_actual += 1

    # CD + CI (merge C:E)
    merge_and_set(ws, f"C{fila_actual}:E{fila_actual}", "CD + CI",
                  font=font_subtotal, fill=fill_gris_seccion, alignment=align_left)
    set_cell(ws, fila_actual, 6, f"=F{cd_row}+F{ci_row}",
             font=font_subtotal, fill=fill_gris_seccion,
             alignment=align_right, number_format="$#,##0.00")
    aplicar_borde(ws, fila_actual, 3, 6, "thin")
    cd_ci_row = fila_actual
    fila_actual += 2

    # Financiamiento (merge C:D, % en E)
    merge_and_set(ws, f"C{fila_actual}:D{fila_actual}", "Financiamiento (F = %(CD+CI))",
                  font=font_normal, alignment=align_left)
    set_cell(ws, fila_actual, 5, FINANCIAMIENTO, font=font_normal, alignment=align_center, number_format="0.00%")
    set_cell(ws, fila_actual, 6, f"=F{cd_ci_row}*E{fila_actual}",
             font=font_normal, alignment=align_right, number_format="$#,##0.00")
    aplicar_borde(ws, fila_actual, 3, 6, "thin")
    fin_row = fila_actual
    fila_actual += 2

    # Utilidad (merge C:D, % en E)
    merge_and_set(ws, f"C{fila_actual}:D{fila_actual}", "Utilidad (U = %(CD+CI+F))",
                  font=font_normal, alignment=align_left)
    set_cell(ws, fila_actual, 5, UTILIDAD, font=font_normal, alignment=align_center, number_format="0%")
    set_cell(ws, fila_actual, 6, f"=(F{cd_ci_row}+F{fin_row})*E{fila_actual}",
             font=font_normal, alignment=align_right, number_format="$#,##0.00")
    aplicar_borde(ws, fila_actual, 3, 6, "thin")
    util_row = fila_actual
    fila_actual += 1

    # PU DE VENTA (merge C:E)
    merge_and_set(ws, f"C{fila_actual}:E{fila_actual}", "PU DE VENTA",
                  font=Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO),
                  fill=PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"),
                  alignment=align_left)
    set_cell(ws, fila_actual, 6, f"=F{cd_ci_row}+F{fin_row}+F{util_row}",
             font=Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO),
             fill=PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"),
             alignment=align_right, number_format="$#,##0.00")
    aplicar_borde(ws, fila_actual, 3, 6, "thin")
    pu_venta_row = fila_actual

    # Actualizar referencia del PU en R3
    set_cell(ws, 3, 6, f"=F{pu_venta_row}",
             font=font_subtotal, alignment=align_right, number_format="$#,##0.00")

    # ---- ANCHOS DE COLUMNA ----
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 22
    ws.column_dimensions["G"].width = 10

    # ---- ALTOS DE FILA ----
    ws.row_dimensions[1].height = 44
    ws.row_dimensions[2].height = 8
    ws.row_dimensions[5].height = 17

    # Proteger celdas con fórmulas (light protection - solo aviso visual)
    # No aplicamos protección real, solo colores para distinguir

    return ws


# =============================================================================
# FUNCION PRINCIPAL
# =============================================================================

def generar_plantilla():
    """Genera la plantilla completa de presupuesto."""
    print(f"\n{'='*60}")
    print(f"  GENERANDO PLANTILLA DE PRESUPUESTO")
    print(f"  Proyecto: {PROYECTO}")
    print(f"  Fecha: {FECHA}")
    print(f"  Conceptos APU: {len(CONCEPTOS_APU)}")
    print(f"  Archivo salida: {ARCHIVO_SALIDA}")
    print(f"{'='*60}\n")

    wb = openpyxl.Workbook()

    # Eliminar hoja default
    wb.remove(wb.active)

    # ---- 1. Hojas de la Calculadora FASAR ----
    print("Creando hoja: Categorias...")
    crear_hoja_categorias(wb)

    print("Creando hoja: Parametros...")
    crear_hoja_parametros(wb)

    print("Creando hoja: Calculo_categoria...")
    crear_hoja_calculo_categoria(wb)

    print("Creando hoja: Calculadora...")
    crear_hoja_calculadora(wb)

    print("Creando hoja: Cuadrillas...")
    crear_hoja_cuadrillas(wb)

    print("Creando hoja: Notas...")
    crear_hoja_notas(wb)

    # ---- 2. Hoja de Variables de Precio ----
    print("Creando hoja: variable...")
    crear_hoja_variables(wb)

    # ---- 3. Hoja de Presupuesto (Resumen/Catálogo) ----
    print("Creando hoja: Presupuesto...")
    crear_hoja_presupuesto(wb, CONCEPTOS_APU)

    # ---- 4. Hojas APU ----
    codigos = []
    for apu in CONCEPTOS_APU:
        codigo = apu["codigo"]
        descripcion = apu["descripcion"]
        unidad = apu["unidad"]
        print(f"Creando hoja APU: {codigo}...")
        crear_hoja_apu(wb, codigo, descripcion, unidad, FILAS_VACIAS_POR_SECCION)
        codigos.append(codigo)

    # ---- 5. Hoja de Insumos (al final para referencia cruzada) ----
    print("Creando hoja: insumos...")
    crear_hoja_insumos(wb, codigos)

    # ---- Guardar ----
    os.makedirs(CARPETA_SALIDA, exist_ok=True)
    print(f"\nGuardando archivo: {ARCHIVO_SALIDA}...")
    wb.save(ARCHIVO_SALIDA)

    # Tamaño del archivo
    size_kb = os.path.getsize(ARCHIVO_SALIDA) / 1024
    print(f"\n{'='*60}")
    print(f"  PLANTILLA GENERADA EXITOSAMENTE")
    print(f"  Archivo: {os.path.abspath(ARCHIVO_SALIDA)}")
    print(f"  Tamaño: {size_kb:.1f} KB")
    print(f"  Hojas: {wb.sheetnames}")
    print(f"{'='*60}\n")

    return ARCHIVO_SALIDA


if __name__ == "__main__":
    generar_plantilla()