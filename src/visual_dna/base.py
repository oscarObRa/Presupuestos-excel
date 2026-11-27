from dataclasses import dataclass
from typing import Optional, Dict, List
from pathlib import Path
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage

@dataclass
class BaseVisualDNA:
    """Clase base con la lógica de estilos, fuentes y formatos comunes."""
    
    # Atributos que cada empresa debe definir
    VERDE_ICEL: str = ""
    VERDE_ENCABEZADO: str = ""
    VERDE_BORDE: str = ""
    GRIS_PARTIDA: str = ""
    
    BLANCO: str = "FFFFFF"
    NEGRO: str = "000000"
    
    # Atributos de tipografía (con valores por defecto)
    FUENTE_TITULO: str = "Arial"
    FUENTE_CUERPO: str = "Arial"
    FUENTE_MONO: str = "Consolas"
    FUENTE_TITULO_FALLBACK: str = "Arial"
    FUENTE_CUERPO_FALLBACK: str = "Calibri"
    
    TAM_TITULO: int = 13
    TAM_SUBTITULO: int = 12
    TAM_ENCABEZADO_TABLA: int = 11
    TAM_CUERPO: int = 10
    TAM_PEQUENO: int = 9
    
    # Assets
    LOGO_PATH: str = "assets/Logo.png"
    FOOTER_PATH: str = "assets/Pie de pagina.png"
    LOGO_ANCHO: int = 120
    FOOTER_ALTO: int = 78
    
    # Margenes
    MARGEN_IZQ: float = 0.5
    MARGEN_DER: float = 0.5
    MARGEN_SUP: float = 0.5
    MARGEN_INF: float = 0.5
    ORIENTACION_PORTADA: str = "portrait"
    ORIENTACION_APU: str = "portrait"
    PAPEL: str = "letter"

    # Cachés
    _font_cache: dict = None
    _fill_cache: dict = None
    _alignment_cache: dict = None
    _border_cache: dict = None

    def __post_init__(self):
        self._font_cache = {}
        self._fill_cache = {}
        self._alignment_cache = {}
        self._border_cache = {}

    # === LOGICA DE ESTILOS (Compartida) ===

    def _get_font(self, **kwargs) -> Font:
        key = tuple(sorted(kwargs.items()))
        if key not in self._font_cache:
            self._font_cache[key] = Font(**kwargs)
        return self._font_cache[key]

    def _get_fill(self, color: str) -> PatternFill:
        if color not in self._fill_cache:
            self._fill_cache[color] = PatternFill(start_color=color, end_color=color, fill_type="solid")
        return self._fill_cache[color]

    def _get_alignment(self, **kwargs) -> Alignment:
        key = tuple(sorted(kwargs.items()))
        if key not in self._alignment_cache:
            self._alignment_cache[key] = Alignment(**kwargs)
        return self._alignment_cache[key]

    def _get_border(self, **kwargs) -> Border:
        key = tuple(sorted(kwargs.items()))
        if key not in self._border_cache:
            self._border_cache[key] = Border(**kwargs)
        return self._border_cache[key]

    # === MÉTODOS DE AYUDA PARA EL CONSTRUCTOR ===

    def font_titulo(self, color: str = None, size: int = None, bold: bool = True) -> Font:
        return self._get_font(name=self.FUENTE_TITULO, size=size or self.TAM_TITULO, bold=bold, color=color or self.NEGRO)

    def font_subtitulo(self, color: str = None, size: int = None, bold: bool = True) -> Font:
        return self._get_font(name=self.FUENTE_TITULO, size=size or self.TAM_SUBTITULO, bold=bold, color=color or self.NEGRO)

    def font_header_tabla(self, blanco: bool = True) -> Font:
        return self._get_font(name=self.FUENTE_TITULO, size=self.TAM_ENCABEZADO_TABLA, bold=True, color=self.BLANCO if blanco else self.NEGRO)

    def font_cuerpo(self, color: str = None, bold: bool = False, size: int = None) -> Font:
        return self._get_font(name=self.FUENTE_CUERPO, size=size or self.TAM_CUERPO, bold=bold, color=color or self.NEGRO)

    def font_partida(self, color: str = None, bold: bool = True, size: int = None) -> Font:
        return self._get_font(name=self.FUENTE_CUERPO, size=size or self.TAM_CUERPO, bold=bold, color=color or self.NEGRO)

    def font_moneda(self, bold: bool = False) -> Font:
        return self._get_font(name=self.FUENTE_CUERPO, size=self.TAM_CUERPO, bold=bold, color=self.NEGRO)

    def font_pequeno(self, color: str = None, bold: bool = False, italic: bool = False) -> Font:
        return self._get_font(name=self.FUENTE_CUERPO, size=self.TAM_PEQUENO, bold=bold,
                              italic=italic, color=color or self.NEGRO)

    # === Variantes fallback (genéricas; cada empresa puede sobrescribirlas) ===

    def font_titulo_fallback(self, color: str = None, size: int = None, bold: bool = True) -> Font:
        return self._get_font(name=self.FUENTE_TITULO_FALLBACK, size=size or self.TAM_TITULO,
                              bold=bold, color=color or self.NEGRO)

    def font_subtitulo_fallback(self, color: str = None, size: int = None, bold: bool = True) -> Font:
        return self._get_font(name=self.FUENTE_TITULO_FALLBACK, size=size or self.TAM_SUBTITULO,
                              bold=bold, color=color or self.NEGRO)

    def font_header_tabla_fallback(self, blanco: bool = True) -> Font:
        return self._get_font(name=self.FUENTE_TITULO_FALLBACK, size=self.TAM_ENCABEZADO_TABLA,
                              bold=True, color=self.BLANCO if blanco else self.NEGRO)

    def font_cuerpo_fallback(self, color: str = None, bold: bool = False, size: int = None) -> Font:
        return self._get_font(name=self.FUENTE_CUERPO_FALLBACK, size=size or self.TAM_CUERPO,
                              bold=bold, color=color or self.NEGRO)

    def fill_verde_icel(self) -> PatternFill: return self._get_fill(self.VERDE_ICEL)
    def fill_verde_encabezado(self) -> PatternFill: return self._get_fill(self.VERDE_ENCABEZADO)
    def fill_verde_profundo(self) -> PatternFill: return self._get_fill(self.VERDE_ICEL) # Reutilización lógica
    def fill_gris_partida(self) -> PatternFill: return self._get_fill(self.GRIS_PARTIDA)
    def fill_blanco(self) -> PatternFill: return self._get_fill(self.BLANCO)
    def fill_negro(self) -> PatternFill: return self._get_fill(self.NEGRO)

    # Colores semánticos de apoyo (heredables por cualquier empresa)
    ROJO_CLARO: str = "FDE9E9"
    VERDE_CLARO_FILL: str = "E2EFDA"
    AMARILLO_CLARO: str = "FFF2CC"
    AZUL_CLARO: str = "DDEBF7"
    NARANJA_CLARO: str = "FCE4D6"

    def fill_rojo_claro(self) -> PatternFill: return self._get_fill(self.ROJO_CLARO)
    def fill_verde_claro(self) -> PatternFill: return self._get_fill(self.VERDE_CLARO_FILL)
    def fill_amarillo_claro(self) -> PatternFill: return self._get_fill(self.AMARILLO_CLARO)
    def fill_azul_claro(self) -> PatternFill: return self._get_fill(self.AZUL_CLARO)
    def fill_naranja_claro(self) -> PatternFill: return self._get_fill(self.NARANJA_CLARO)

    def align_center(self) -> Alignment: return self._get_alignment(horizontal="center", vertical="center", wrap_text=True)
    def align_left(self) -> Alignment: return self._get_alignment(horizontal="left", vertical="center", wrap_text=True)
    def align_right(self) -> Alignment: return self._get_alignment(horizontal="right", vertical="center")
    def align_center_nowrap(self) -> Alignment: return self._get_alignment(horizontal="center", vertical="center")

    def border_header_tabla(self) -> Border:
        """Bordes del encabezado: exterior medium, interiores thin (ADN)."""
        thin = Side(style="thin", color=self._color_borde())
        medium = Side(style="medium", color=self._color_borde())
        return self._get_border(top=medium, bottom=thin, left=thin, right=thin)

    def border_celda(self) -> Border:
        """Borde estándar de celda de datos: thin en los 4 lados (ADN)."""
        thin = Side(style="thin", color=self._color_borde())
        return self._get_border(top=thin, bottom=thin, left=thin, right=thin)

    def border_partida(self) -> Border:
        """Borde de títulos de partida: medium arriba, thin abajo."""
        thin = Side(style="thin", color=self._color_borde())
        medium = Side(style="medium", color=self._color_borde())
        return self._get_border(top=medium, bottom=thin, left=thin, right=thin)

    def _color_borde(self) -> str:
        """Color de borde del ADN; fallback negro si la empresa no lo define."""
        return self.VERDE_BORDE or self.NEGRO

    def configurar_hoja_limpia(self, ws):
        """Vista limpia sin cuadrícula; aplica la retícula del ADN si la define."""
        ws.views.sheetView[0].showGridLines = False
        anchos = getattr(self, "COLUMN_WIDTHS", None)
        if anchos:
            for col, ancho in anchos.items():
                ws.column_dimensions[col].width = ancho
        
    def configurar_impresion(self, ws):
        ws.page_setup.orientation = self.ORIENTACION_PORTADA
        ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = self.MARGEN_IZQ
        ws.page_margins.right = self.MARGEN_DER
        ws.page_margins.top = self.MARGEN_SUP
        ws.page_margins.bottom = self.MARGEN_INF
