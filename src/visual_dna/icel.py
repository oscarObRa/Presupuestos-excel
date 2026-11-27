"""
icel.py
=======
ADN Visual de ICEL Contratistas S.A. de C.V.

Fuente de verdad: docs/ADN_ICEL.md
- Paleta institucional (HEX sin '#', inmune al tema de Office)
- Retícula de 17 columnas (A:Q) y mapa de coordenadas de la carátula
- Metadatos fiscales de la empresa (domicilio, RFC, assets)

Para agregar otra empresa: crear un módulo hermano (ej. otra_empresa.py)
con una clase que herede de BaseVisualDNA y registrarla en factory.py.
"""

from dataclasses import dataclass

from src.visual_dna.base import BaseVisualDNA

# === PALETA INSTITUCIONAL (RGB explícito, sin '#') ===
FILL_TABLE_HEADER = "3C7D22"    # Verde oscuro para cabecera de tabla
COLOR_SECTION_ACCENT = "4EA72E" # Verde corporativo para títulos de sección y acentos
COLOR_BORDER_DARK = "275317"    # Verde profundo para bordes de tabla
FILL_PART_ROW = "F2F5F0"        # Fondo suave para filas de partida

FONT_TEXT_WHITE = "FFFFFF"
FONT_TEXT_BLACK = "000000"

# === METADATOS DE LA EMPRESA (membrete) ===
EMPRESA_NOMBRE = "ICEL Contratistas S.A. de C.V."
EMPRESA_RFC = "ICO200226NP1"
EMPRESA_DOMICILIO = [
    "Vía Lombardía 5701, Suite No. 409",
    "Edificio Corporativo Distrito 1",
    "Chihuahua, Chih.",
]

# === RETÍCULA DE CARÁTULA (17 columnas A:Q) ===
GRID_COLUMNS = [
    "A", "B", "C", "D", "E", "F", "G", "H", "I",
    "J", "K", "L", "M", "N", "O", "P", "Q",
]

COLUMN_WIDTHS = {
    "A": 3.8,  "B": 13.0, "C": 4.6,  "D": 6.4,  "E": 4.2,  "F": 1.5,
    "G": 4.0,  "H": 3.5,  "I": 12.2, "J": 4.5,  "K": 72.8, "L": 11.5,
    "M": 11.8, "N": 18.0, "O": 7.5,  "P": 7.2,  "Q": 4.6,
}

# Mapa de coordenadas fijas de la carátula / propuesta
CELL_MAP = {
    "lugar_fecha": "A7",
    "folio": "L7",
    "etiqueta_propuesta": "A10",
    "nombre_propuesta": "D10",
    "etiqueta_atencion": "A13",
    "destinatario": "C13",
    "cargo": "C14",
    "empresa": "C15",
    "introduccion": "C17",
    "seccion_economica": "A20",
    "fila_encabezado_tabla": 22,
    "fila_inicio_conceptos": 23,
}

DEFAULT_FONT_FAMILY = "Arial"


@dataclass
class ICELVisualDNA(BaseVisualDNA):
    """Identidad visual ICEL Contratistas."""

    # --- Paleta ---
    VERDE_ICEL: str = COLOR_SECTION_ACCENT     # Secciones, acentos, membrete
    VERDE_ENCABEZADO: str = FILL_TABLE_HEADER  # Relleno de headers de tabla
    VERDE_BORDE: str = COLOR_BORDER_DARK       # Bordes de tabla
    GRIS_PARTIDA: str = FILL_PART_ROW          # Fondo de filas de partida
    BLANCO: str = FONT_TEXT_WHITE
    NEGRO: str = FONT_TEXT_BLACK

    # Colores de apoyo (semánticos)
    ROJO_CLARO: str = "FDE9E9"
    VERDE_CLARO_FILL: str = "E2EFDA"
    AMARILLO_CLARO: str = "FFF2CC"
    AZUL_CLARO: str = "DDEBF7"
    NARANJA_CLARO: str = "FCE4D6"

    # --- Tipografía (estandarizada a Arial para exportación PDF estable) ---
    FUENTE_TITULO: str = DEFAULT_FONT_FAMILY
    FUENTE_CUERPO: str = DEFAULT_FONT_FAMILY
    FUENTE_TITULO_FALLBACK: str = DEFAULT_FONT_FAMILY

    # --- Assets propios de ICEL ---
    LOGO_PATH: str = "assets/Logo.png"
    FOOTER_PATH: str = "assets/Pie de pagina.png"

    # --- Metadatos para el membrete ---
    NOMBRE_EMPRESA: str = EMPRESA_NOMBRE
    RFC: str = EMPRESA_RFC
    DOMICILIO: list = None  # se completa en __post_init__ (lista mutable)

    # --- Retícula ---
    GRID_COLUMNS: list = None
    COLUMN_WIDTHS: dict = None
    CELL_MAP: dict = None

    def __post_init__(self):
        super().__post_init__()
        self.DOMICILIO = list(EMPRESA_DOMICILIO)
        self.GRID_COLUMNS = list(GRID_COLUMNS)
        self.COLUMN_WIDTHS = dict(COLUMN_WIDTHS)
        self.CELL_MAP = dict(CELL_MAP)

    # === Fuentes específicas ICEL ===

    def font_cuerpo_fallback(self, color: str = None, bold: bool = False, size: int = None):
        """En ICEL, el fallback es la misma Arial: un solo estilo."""
        return self.font_cuerpo(color=color, bold=bold, size=size)

    def font_subtitulo_fallback(self, color: str = None, size: int = None, bold: bool = True):
        return self.font_subtitulo(color=color, size=size, bold=bold)

    def font_header_tabla_fallback(self, blanco: bool = True):
        return self.font_header_tabla(blanco=blanco)

    def font_titulo_fallback(self, color: str = None, size: int = None, bold: bool = True):
        return self.font_titulo(color=color, size=size, bold=bold)
