"""Estilos, colores y formatos para la plantilla de presupuesto.

Basado en el Catálogo muro rev 2.
Uso:
    from src.estilos import (
        AZUL_OSCURO, AZUL_HEADER, GRIS_SECCION,
        fill_azul_oscuro, fill_azul_header, fill_gris_seccion,
        font_titulo, font_header_blanco, font_normal, font_seccion,
        align_center, align_left, align_right,
        borde_fino, borde_medio_left,
        aplicar_borde, merge_y_escribir, escribir_celda,
    )
"""

from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side,
)

# =============================================================================
# COLORES
# =============================================================================

AZUL_OSCURO = "1F3864"
AZUL_HEADER = "DDEBF7"
AZUL_MEDIO = "1F4E78"
GRIS_SECCION = "F2F2F2"
BLANCO = "FFFFFF"
NEGRO_TEXTO = "000000"
ROJO_TEXTO = "FF0000"
VERDE_CLARO = "E2F0D9"
AMARILLO_CLARO = "FFF2CC"
AZUL_PU_VENTA = "D9E1F2"

# =============================================================================
# FILLS
# =============================================================================

fill_azul_oscuro = PatternFill(start_color=AZUL_OSCURO, end_color=AZUL_OSCURO, fill_type="solid")
fill_azul_header = PatternFill(start_color=AZUL_HEADER, end_color=AZUL_HEADER, fill_type="solid")
fill_azul_medio = PatternFill(start_color=AZUL_MEDIO, end_color=AZUL_MEDIO, fill_type="solid")
fill_gris_seccion = PatternFill(start_color=GRIS_SECCION, end_color=GRIS_SECCION, fill_type="solid")
fill_blanco = PatternFill(start_color=BLANCO, end_color=BLANCO, fill_type="solid")
fill_verde_claro = PatternFill(start_color=VERDE_CLARO, end_color=VERDE_CLARO, fill_type="solid")
fill_amarillo_claro = PatternFill(start_color=AMARILLO_CLARO, end_color=AMARILLO_CLARO, fill_type="solid")
fill_pu_venta = PatternFill(start_color=AZUL_PU_VENTA, end_color=AZUL_PU_VENTA, fill_type="solid")
fill_negro = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
fill_azul_presupuesto = PatternFill(start_color="003366", end_color="003366", fill_type="solid")

# =============================================================================
# FONTS
# =============================================================================

font_titulo = Font(name="Calibri", size=14, bold=True, color=NEGRO_TEXTO)
font_titulo_blanco = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
font_subtitulo = Font(name="Calibri", size=12, bold=True, color=NEGRO_TEXTO)
font_header_blanco = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
font_header_negro = Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO)
font_header_col = Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO)
font_seccion = Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO)
font_normal = Font(name="Calibri", size=10, color=NEGRO_TEXTO)
font_normal_rojo = Font(name="Calibri", size=10, color=ROJO_TEXTO)
font_negrita = Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO)
font_italica = Font(name="Calibri", size=9, italic=True, color="404040")
font_apu_titulo = Font(name="Calibri", size=11, bold=True, color=NEGRO_TEXTO)
font_apu_concepto = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
font_subtotal = Font(name="Calibri", size=10, bold=True, color=NEGRO_TEXTO)
font_param_header = Font(name="Calibri", size=11, bold=True, color=NEGRO_TEXTO)
font_valor_editable = Font(name="Calibri", size=10, color=NEGRO_TEXTO)

# =============================================================================
# ALIGNMENTS
# =============================================================================

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_right = Alignment(horizontal="right", vertical="center")
align_center_nowrap = Alignment(horizontal="center", vertical="center")

# =============================================================================
# BORDERS
# =============================================================================

borde_fino = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
borde_medio_left = Border(
    left=Side(style="medium"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
borde_inferior = Border(bottom=Side(style="thin"))
borde_superior_grueso = Border(top=Side(style="thick"))
borde_doble_inferior = Border(bottom=Side(style="double"))

# =============================================================================
# FUNCIONES AUXILIARES
# =============================================================================


def aplicar_borde(ws, fila, col_inicio, col_fin, estilo="thin"):
    """Aplica bordes a un rango horizontal de celdas."""
    for col in range(col_inicio, col_fin + 1):
        cell = ws.cell(row=fila, column=col)
        cell.border = Border(
            left=Side(style=estilo), right=Side(style=estilo),
            top=Side(style=estilo), bottom=Side(style=estilo),
        )


def merge_y_escribir(ws, rango, valor, font=None, fill=None, alignment=None):
    """Hace merge de un rango y asigna valor + estilos a la celda principal."""
    ws.merge_cells(rango)
    celda = ws[rango.split(":")[0]]
    celda.value = valor
    if font:
        celda.font = font
    if fill:
        celda.fill = fill
    if alignment:
        celda.alignment = alignment
    return celda


def escribir_celda(ws, fila, col, valor, font=None, fill=None, alignment=None,
                   number_format=None, border=None):
    """Asigna valor y estilos a una celda individual."""
    celda = ws.cell(row=fila, column=col)
    celda.value = valor
    if font:
        celda.font = font
    if fill:
        celda.fill = fill
    if alignment:
        celda.alignment = alignment
    if number_format:
        celda.number_format = number_format
    if border:
        celda.border = border
    return celda


def col_letra(n):
    """Convierte número de columna a letra (1→A, 2→B, ..., 27→AA)."""
    letra = ""
    while n > 0:
        n, resto = divmod(n - 1, 26)
        letra = chr(65 + resto) + letra
    return letra