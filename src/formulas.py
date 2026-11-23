"""
formulas.py
===========
Helpers para generar fórmulas nativas de Excel con openpyxl.

Uso:
    from src.formulas import (
        VLOOKUP, SUMIF, SUMPRODUCT, ref_hoja, celda, rango,
        col_letra, formula_suma, formula_multiplicacion,
    )

    # =VLOOKUP(A5, Insumos!A:J, 3, FALSE)
    formula = VLOOKUP("A5", "Insumos", "A:J", 3)

    # =SUMIF(A8:A50, "PRELIMINARES", G8:G50)
    formula = SUMIF("A8:A50", "PRELIMINARES", "G8:G50")

    # ='Calculo_categoria'!K5
    formula = ref_hoja("Calculo_categoria", "K5")
"""

import re


def col_letra(n: int) -> str:
    """Convierte número de columna a letra: 1→A, 2→B, ..., 27→AA."""
    letra = ""
    while n > 0:
        n, resto = divmod(n - 1, 26)
        letra = chr(65 + resto) + letra
    return letra


def celda(col: int, fila: int) -> str:
    """Retorna referencia Excel: celda(3, 5) → 'C5'."""
    return f"{col_letra(col)}{fila}"


def rango(col_ini: int, fila_ini: int, col_fin: int, fila_fin: int) -> str:
    """Retorna rango: rango(1, 5, 3, 10) → 'A5:C10'."""
    return f"{celda(col_ini, fila_ini)}:{celda(col_fin, fila_fin)}"


# ---------------------------------------------------------------------------
# Referencias entre hojas
# ---------------------------------------------------------------------------

def _escapar_hoja(nombre: str) -> str:
    """Escapa nombres de hoja con espacios o caracteres especiales.

    'Mi Hoja' → "'Mi Hoja'"
    'Insumos' → 'Insumos'
    """
    if re.search(r"[ '\-]", nombre):
        return f"'{nombre}'"
    return nombre


def ref_hoja(hoja: str, ref: str) -> str:
    """Referencia a celda/rango en otra hoja.

    ref_hoja("Insumos", "A5")      → 'Insumos!A5'
    ref_hoja("Mi Hoja", "B10")     → "'Mi Hoja'!B10"
    ref_hoja("Calculo_categoria", "K5:K50") → "'Calculo_categoria'!K5:K50"
    """
    return f"{_escapar_hoja(hoja)}!{ref}"


# ---------------------------------------------------------------------------
# Fórmulas comunes
# ---------------------------------------------------------------------------

def VLOOKUP(lookup: str, hoja: str, rango_ref: str, col: int,
            exacto: bool = True) -> str:
    """Genera fórmula VLOOKUP.

    VLOOKUP("A5", "Insumos", "A:J", 3) → '=VLOOKUP(A5, Insumos!A:J, 3, FALSE)'

    Args:
        lookup: Celda con el valor a buscar (ej. "A5").
        hoja: Nombre de la hoja donde buscar.
        rango_ref: Rango de búsqueda (ej. "A:J" o "A2:J500").
        col: Número de columna (1-based) de donde retornar el valor.
        exacto: True → FALSE (búsqueda exacta); False → TRUE (aproximada).
    """
    match_type = "FALSE" if exacto else "TRUE"
    return f"=VLOOKUP({lookup}, {_escapar_hoja(hoja)}!{rango_ref}, {col}, {match_type})"


def VLOOKUP_col(lookup_col: int, lookup_fila: int, hoja: str,
                rango_ref: str, col_retorno: int, exacto: bool = True) -> str:
    """VLOOKUP usando número de columna/fila en vez de string.

    VLOOKUP_col(1, 5, "Insumos", "A:J", 3) → '=VLOOKUP(A5, Insumos!A:J, 3, FALSE)'
    """
    lookup = celda(lookup_col, lookup_fila)
    return VLOOKUP(lookup, hoja, rango_ref, col_retorno, exacto)


def SUMIF(rango_criterio: str, criterio: str, rango_suma: str) -> str:
    """Genera fórmula SUMIF.

    SUMIF("A8:A50", "PRELIMINARES", "G8:G50")
    → '=SUMIF(A8:A50, "PRELIMINARES", G8:G50)'

    SUMIF("A8:A50", "A8", "G8:G50")
    → '=SUMIF(A8:A50, A8, G8:G50)'
    """
    # Si el criterio no empieza con comilla y no es referencia de celda,
    # lo envolvemos en comillas para Excel
    if not criterio.startswith('"') and not re.match(r'^[A-Z]+\d+$', criterio):
        criterio = f'"{criterio}"'
    return f"=SUMIF({rango_criterio}, {criterio}, {rango_suma})"


def SUMPRODUCT(rango1: str, rango2: str) -> str:
    """Genera fórmula SUMPRODUCT.

    SUMPRODUCT("E8:E50", "F8:F50") → '=SUMPRODUCT(E8:E50, F8:F50)'
    """
    return f"=SUMPRODUCT({rango1}, {rango2})"


def SUM(rango_ref: str) -> str:
    """Genera fórmula SUM."""
    return f"=SUM({rango_ref})"


def IF(condicion: str, valor_si: str, valor_no: str) -> str:
    """Genera fórmula IF."""
    return f"=IF({condicion}, {valor_si}, {valor_no})"


def MIN(a: str, b: str) -> str:
    """Genera fórmula MIN."""
    return f"=MIN({a}, {b})"


# ---------------------------------------------------------------------------
# Construcción de fórmulas compuestas
# ---------------------------------------------------------------------------

def formula_producto(celda_cant: str, celda_pu: str) -> str:
    """Cantidad × Precio Unitario."""
    return f"={celda_cant}*{celda_pu}"


def formula_porcentaje(celda_base: str, celda_pct: str) -> str:
    """Base × Porcentaje."""
    return f"={celda_base}*{celda_pct}"


def formula_total_con_indirectos(cd: str, ci: str, fin: str, util: str) -> str:
    """PU de venta = CD + CI + Fin + Util."""
    return f"={cd}+{ci}+{fin}+{util}"


def formula_indirectos(celda_cd: str, celda_pct_ind: str) -> str:
    """Indirectos = CD × %Indirectos."""
    return f"={celda_cd}*{celda_pct_ind}"


def formula_utilidad(celda_base: str, celda_pct_util: str) -> str:
    """Utilidad = (CD + CI + Fin) × %Utilidad."""
    return f"=({celda_base})*{celda_pct_util}"


# ---------------------------------------------------------------------------
# Test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("=== Helpers de fórmulas ===\n")

    print("col_letra(1):", col_letra(1))        # A
    print("col_letra(27):", col_letra(27))      # AA
    print("celda(3, 5):", celda(3, 5))           # C5
    print("rango(1, 5, 3, 10):", rango(1, 5, 3, 10))  # A5:C10

    print()
    print("VLOOKUP:", VLOOKUP("A5", "Insumos", "A:J", 3))
    print("VLOOKUP con espacios:", VLOOKUP("B10", "Mi Hoja", "A:K", 11))
    print("SUMIF:", SUMIF("A8:A50", "PRELIMINARES", "G8:G50"))
    print("SUMIF por celda:", SUMIF("A8:A50", "A8", "G8:G50"))
    print("SUMPRODUCT:", SUMPRODUCT("E8:E50", "F8:F50"))

    print()
    print("Total con indirectos:", formula_total_con_indirectos("F40", "F42", "F44", "F45"))
    print("ref_hoja:", ref_hoja("Calculo_categoria", "K5"))

    print("\n=== TODO OK ===")