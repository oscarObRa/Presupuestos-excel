"""
visual_dna (paquete)
====================
ADN Visual del sistema — identidad visual por empresa.

Estructura:
- base.py     : BaseVisualDNA (lógica común de fuentes/fills/bordes/impresión)
- icel.py     : ICELVisualDNA (paleta, retícula A:Q y metadatos de ICEL)
- factory.py  : get_dna(empresa) → instancia del ADN correspondiente

Uso en generadores:
    from src.visual_dna import get_dna
    dna = get_dna(proyecto.get("empresa"))   # None → empresa por defecto

Compatibilidad: se mantiene `DNA` (instancia ICEL) y la clase `VisualDNA`
para que `from src.visual_dna import DNA` siga funcionando mientras se
migran los usos restantes.
"""

from src.visual_dna.base import BaseVisualDNA
from src.visual_dna.icel import (
    ICELVisualDNA,
    FILL_TABLE_HEADER,
    COLOR_SECTION_ACCENT,
    COLOR_BORDER_DARK,
    FILL_PART_ROW,
    FONT_TEXT_WHITE,
    FONT_TEXT_BLACK,
    COLUMN_WIDTHS,
    CELL_MAP,
    GRID_COLUMNS,
    DEFAULT_FONT_FAMILY,
)
from src.visual_dna.factory import VisualDNAFactory, get_dna

# --- Compatibilidad con el contrato anterior (src/visual_dna.py) ---
VisualDNA = ICELVisualDNA
DNA = get_dna()  # instancia ICEL por defecto, igual que antes

__all__ = [
    "BaseVisualDNA",
    "VisualDNA",
    "DNA",
    "ICELVisualDNA",
    "VisualDNAFactory",
    "get_dna",
    # Paleta ICEL
    "FILL_TABLE_HEADER",
    "COLOR_SECTION_ACCENT",
    "COLOR_BORDER_DARK",
    "FILL_PART_ROW",
    "FONT_TEXT_WHITE",
    "FONT_TEXT_BLACK",
    # Retícula
    "COLUMN_WIDTHS",
    "CELL_MAP",
    "GRID_COLUMNS",
    "DEFAULT_FONT_FAMILY",
]
