"""
visual_dna.py
=============
ADN Visual del sistema - Define la identidad visual base (ICEL Contratistas)
que se usa como base para todos los documentos generados.

No hardcodear colores: todo viene de aquí.

Colores extraídos del Pie de pagina.png y caratula.xlsx (ICEL Contratistas):
- Verde Olivo ICEL: #4E623A (footer principal)
- Verde Bosque Office: #3B7D23 (encabezados tabla actual)
- Verde Profundo: #275417 (subtítulos/acentos oscuros)
- Gris Antracita: #484B50 (logo/acentos ICEL)
- Gris Medio: #A6A6A6 (franjas diagonales footer)
- Blanco: #FFFFFF
- Negro: #000000
"""

from dataclasses import dataclass
from typing import Optional
from pathlib import Path
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage


@dataclass
class VisualDNA:
    """ADN Visual de la marca ICEL - colores, tipografías, logo, footer."""

    # === PALETA DE COLORES ICEL (HEX fijos, inmunes a cambios de tema) ===
    # Colores corporativos principales (extraídos de Pie de pagina.png)
    VERDE_ICEL: str = "4E623A"          # Verde Olivo ICEL - color principal corporativo
    VERDE_BOSQUE: str = "3B7D23"        # Verde Bosque Office - encabezados de tabla actuales
    VERDE_PROFUNDO: str = "275417"      # Verde Profundo - subtítulos, acentos oscuros
    GRIS_ANTRACITA: str = "484B50"      # Gris Antracita - logo ICEL, acentos
    GRIS_MEDIO: str = "A6A6A6"          # Gris Medio - franjas diagonales footer
    
    # Colores de apoyo
    BLANCO: str = "FFFFFF"
    NEGRO: str = "000000"
    
    # Colores semánticos para UI (rojo/verde/amarillo estándar)
    ROJO_CLARO: str = "FDE9E9"          # Para revisión/dudosos
    VERDE_CLARO_FILL: str = "E2EFDA"    # Totales, confirmados
    AMARILLO_CLARO: str = "FFF2CC"      # Advertencias, celdas editables
    AZUL_CLARO: str = "DDEBF7"          # Headers de tabla estilo anterior (legacy)
    NARANJA_CLARO: str = "FCE4D6"       # Secciones APU (Materiales, MO, etc.)

    # === TIPOGRAFÍA ===
    # AvenirNext LT Pro Regular (con respaldo Arial / Calibri)
    FUENTE_TITULO: str = "AvenirNext LT Pro Regular"
    FUENTE_CUERPO: str = "AvenirNext LT Pro Regular"
    FUENTE_MONO: str = "Consolas"
    
    # Fallback sources si la fuente no está instalada
    FUENTE_TITULO_FALLBACK: str = "Arial"
    FUENTE_CUERPO_FALLBACK: str = "Calibri"

    TAM_TITULO: int = 24          # Razón social / Título principal
    TAM_SUBTITULO: int = 12       # Títulos de sección (I.- Propuesta Económica)
    TAM_ENCABEZADO_TABLA: int = 12  # Encabezados de tabla (fila 22)
    TAM_CUERPO: int = 11          # Conceptos y precios
    TAM_PEQUENO: int = 9          # Notas, cláusulas
    
    # === LOGO Y FOOTER ===
    LOGO_PATH: str = "assets/Logo.png"
    FOOTER_PATH: str = "assets/Pie de pagina.png"
    LOGO_ANCHO: int = 120
    FOOTER_ALTO: int = 78
    
    # === MÁRGENES E IMPRESIÓN ===
    MARGEN_IZQ: float = 0.5
    MARGEN_DER: float = 0.5
    MARGEN_SUP: float = 0.5
    MARGEN_INF: float = 0.5
    ORIENTACION_PORTADA: str = "portrait"
    ORIENTACION_APU: str = "portrait"
    PAPEL: str = "letter"  # PAPERSIZE_LETTER
    
    # === ANCHOS DE COLUMNA ESTÁNDAR (carátula propuesta) ===
    COL_PART: float = 8
    COL_CONCEPTO: float = 65
    COL_UNIDAD: float = 12
    COL_CANTIDAD: float = 12
    COL_PRECIO_U: float = 16
    COL_IMPORTE: float = 18

    # === CACHÉ DE ESTILOS ===
    _font_cache: dict = None
    _fill_cache: dict = None
    _alignment_cache: dict = None
    _border_cache: dict = None
    
    def __post_init__(self):
        self._font_cache = {}
        self._fill_cache = {}
        self._alignment_cache = {}
        self._border_cache = {}

    # === FUENTES ===
    
    def _get_font(self, **kwargs) -> Font:
        key = tuple(sorted(kwargs.items()))
        if key not in self._font_cache:
            self._font_cache[key] = Font(**kwargs)
        return self._font_cache[key]

    def font_titulo(self, color: str = None, size: int = None, bold: bool = True) -> Font:
        """Título principal (Razón social) - 24pt, negrita."""
        return self._get_font(
            name=self.FUENTE_TITULO,
            size=size or self.TAM_TITULO,
            bold=bold,
            color=color or self.NEGRO
        )

    def font_titulo_fallback(self, color: str = None, size: int = None, bold: bool = True) -> Font:
        return self._get_font(
            name=self.FUENTE_TITULO_FALLBACK,
            size=size or self.TAM_TITULO,
            bold=bold,
            color=color or self.NEGRO
        )

    def font_subtitulo(self, color: str = None, size: int = None, bold: bool = True) -> Font:
        """Títulos de sección (I.- Propuesta Económica) - 12pt, negrita."""
        return self._get_font(
            name=self.FUENTE_TITULO,
            size=size or self.TAM_SUBTITULO,
            bold=bold,
            color=color or self.NEGRO
        )

    def font_subtitulo_fallback(self, color: str = None, size: int = None, bold: bool = True) -> Font:
        return self._get_font(
            name=self.FUENTE_TITULO_FALLBACK,
            size=size or self.TAM_SUBTITULO,
            bold=bold,
            color=color or self.NEGRO
        )

    def font_header_tabla(self, blanco: bool = True) -> Font:
        """Encabezados de tabla (fila 22) - 12pt, negrita, texto blanco/negro."""
        return self._get_font(
            name=self.FUENTE_TITULO,
            size=self.TAM_ENCABEZADO_TABLA,
            bold=True,
            color=self.BLANCO if blanco else self.NEGRO
        )

    def font_header_tabla_fallback(self, blanco: bool = True) -> Font:
        return self._get_font(
            name=self.FUENTE_TITULO_FALLBACK,
            size=self.TAM_ENCABEZADO_TABLA,
            bold=True,
            color=self.BLANCO if blanco else self.NEGRO
        )

    def font_partida(self, color: str = None, bold: bool = True, size: int = None) -> Font:
        """Títulos de partida/capítulo - 11-12pt, negrita."""
        return self._get_font(
            name=self.FUENTE_CUERPO,
            size=size or self.TAM_CUERPO,
            bold=bold,
            color=color or self.NEGRO
        )

    def font_cuerpo(self, color: str = None, bold: bool = False, size: int = None) -> Font:
        """Texto de cuerpo - 11pt, regular."""
        return self._get_font(
            name=self.FUENTE_CUERPO,
            size=size or self.TAM_CUERPO,
            bold=bold,
            color=color or self.NEGRO
        )

    def font_cuerpo_fallback(self, color: str = None, bold: bool = False, size: int = None) -> Font:
        return self._get_font(
            name=self.FUENTE_CUERPO_FALLBACK,
            size=size or self.TAM_CUERPO,
            bold=bold,
            color=color or self.NEGRO
        )

    def font_pequeno(self, color: str = None, bold: bool = False, italic: bool = False) -> Font:
        """Notas, cláusulas - 9pt."""
        return self._get_font(
            name=self.FUENTE_CUERPO,
            size=self.TAM_PEQUENO,
            bold=bold,
            italic=italic,
            color=color or self.NEGRO
        )

    def font_moneda(self, bold: bool = False) -> Font:
        """Valores monetarios - alineado derecha, formato $."""
        return self._get_font(
            name=self.FUENTE_CUERPO,
            size=self.TAM_CUERPO,
            bold=bold,
            color=self.NEGRO
        )

    # === FILLS (rellenos) ===
    
    def _get_fill(self, color: str) -> PatternFill:
        if color not in self._fill_cache:
            self._fill_cache[color] = PatternFill(start_color=color, end_color=color, fill_type="solid")
        return self._fill_cache[color]

    def fill_verde_icel(self) -> PatternFill:
        """Verde Olivo ICEL - Header principal, footer."""
        return self._get_fill(self.VERDE_ICEL)

    def fill_verde_bosque(self) -> PatternFill:
        """Verde Bosque Office - Encabezados de tabla."""
        return self._get_fill(self.VERDE_BOSQUE)

    def fill_verde_profundo(self) -> PatternFill:
        """Verde Profundo - Subtítulos, acentos."""
        return self._get_fill(self.VERDE_PROFUNDO)

    def fill_gris_antracita(self) -> PatternFill:
        """Gris Antracita - Logo, acentos."""
        return self._get_fill(self.GRIS_ANTRACITA)

    def fill_gris_medio(self) -> PatternFill:
        """Gris Medio - Franjas diagonales."""
        return self._get_fill(self.GRIS_MEDIO)

    def fill_blanco(self) -> PatternFill:
        return self._get_fill(self.BLANCO)

    def fill_negro(self) -> PatternFill:
        return self._get_fill(self.NEGRO)

    def fill_rojo_claro(self) -> PatternFill:
        return self._get_fill(self.ROJO_CLARO)

    def fill_verde_claro(self) -> PatternFill:
        return self._get_fill(self.VERDE_CLARO_FILL)

    def fill_amarillo_claro(self) -> PatternFill:
        return self._get_fill(self.AMARILLO_CLARO)

    def fill_azul_claro(self) -> PatternFill:
        return self._get_fill(self.AZUL_CLARO)

    def fill_naranja_claro(self) -> PatternFill:
        return self._get_fill(self.NARANJA_CLARO)

    # === ALIGNMENTS ===
    
    def _get_alignment(self, **kwargs) -> Alignment:
        key = tuple(sorted(kwargs.items()))
        if key not in self._alignment_cache:
            self._alignment_cache[key] = Alignment(**kwargs)
        return self._alignment_cache[key]

    def align_center(self) -> Alignment:
        return self._get_alignment(horizontal="center", vertical="center", wrap_text=True)

    def align_center_nowrap(self) -> Alignment:
        return self._get_alignment(horizontal="center", vertical="center")

    def align_left(self) -> Alignment:
        return self._get_alignment(horizontal="left", vertical="center", wrap_text=True)

    def align_right(self) -> Alignment:
        return self._get_alignment(horizontal="right", vertical="center")

    # === BORDERS ===
    
    def _get_border(self, **kwargs) -> Border:
        key = tuple(sorted(kwargs.items()))
        if key not in self._border_cache:
            self._border_cache[key] = Border(**kwargs)
        return self._border_cache[key]

    def border_header_tabla(self) -> Border:
        """Borde para encabezado de tabla: medio arriba, fino abajo, fino lados."""
        thin = Side(style="thin", color=self.NEGRO)
        medium = Side(style="medium", color=self.NEGRO)
        return self._get_border(top=medium, bottom=thin, left=thin, right=thin)

    def border_celda(self) -> Border:
        """Borde estándar para celdas de datos: fino en los 4 lados."""
        thin = Side(style="thin", color=self.NEGRO)
        return self._get_border(top=thin, bottom=thin, left=thin, right=thin)

    def border_partida(self) -> Border:
        """Borde para títulos de partida: medio arriba, fino abajo."""
        thin = Side(style="thin", color=self.NEGRO)
        medium = Side(style="medium", color=self.NEGRO)
        return self._get_border(top=medium, bottom=thin, left=thin, right=thin)

    # === HELPERS ===
    
    def configurar_hoja_limpia(self, ws):
        """Configura vista limpia sin líneas de cuadrícula."""
        ws.views.sheetView[0].showGridLines = False
        
        # Anchos de columna estándar para propuesta
        ws.column_dimensions['A'].width = self.COL_PART
        ws.column_dimensions['B'].width = self.COL_CONCEPTO
        ws.column_dimensions['C'].width = self.COL_UNIDAD
        ws.column_dimensions['D'].width = self.COL_CANTIDAD
        ws.column_dimensions['E'].width = self.COL_PRECIO_U
        ws.column_dimensions['F'].width = self.COL_IMPORTE

    def configurar_impresion(self, ws):
        """Configura página para exportar a PDF (ajustar a 1 página ancho)."""
        ws.page_setup.orientation = self.ORIENTACION_PORTADA
        ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0  # 0 = automático
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = self.MARGEN_IZQ
        ws.page_margins.right = self.MARGEN_DER
        ws.page_margins.top = self.MARGEN_SUP
        ws.page_margins.bottom = self.MARGEN_INF


# Instancia global por conveniencia
DNA = VisualDNA()