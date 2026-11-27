"""
lector_entrada.py
=================
Lee el archivo XLSX de entrada con los conceptos a cotizar.

Formatos soportados:
- 4 columnas: id | concepto | unidad | cantidad
- 5 columnas: id | partida | concepto | unidad | cantidad

Además, el formato de 4 columnas puede contener filas de PARTIDA y SUBPARTIDA
incrustadas: filas donde la columna 'id' empieza con "PARTIDA" o "SUB" y las
columnas restantes están vacías. Estas filas agrupan los conceptos siguientes.

Ejemplo:
    PARTIDA 01: CIMENTACIÓN | (vacío) | (vacío) | (vacío)
    CIM-01                  | Excavación... | m3 | 9.75
    CIM-02                  | Relleno...    | m3 | 0.50
    PARTIDA 02: MUROS ...
    MPB-01                  | Castillos...  | ml | 54.00

Uso:
    from src.lector_entrada import leer_entrada

    conceptos = leer_entrada("entrada/conceptos.xlsx")
    for c in conceptos:
        print(c["partida"], c["id"], c["concepto"])
"""

import re
from pathlib import Path
from typing import Optional

import pandas as pd


# Columnas esperadas en orden (ambos formatos)
COLUMNAS_4 = ["id", "concepto", "unidad", "cantidad"]
COLUMNAS_5 = ["id", "partida", "concepto", "unidad", "cantidad"]

# Patrones que identifican una fila de partida/subpartida
PATRON_PARTIDA = re.compile(r"^\s*(PARTIDA|SUBPARTIDA|SUB-?PARTIDA)", re.IGNORECASE)


def _columna_vacia(serie: pd.Series) -> bool:
    """True si la serie no tiene ningún valor con contenido real."""
    vals = serie.astype(str).str.strip().str.lower()
    return bool(vals.isin(["", "nan", "none"]).all())


def _descartar_columnas_fantasma(df: pd.DataFrame) -> pd.DataFrame:
    """Elimina columnas 'Unnamed: N' totalmente vacías.

    El Excel de entrada suele traer celdas con formato (bordes, colores) más
    allá de la última columna de datos: pandas las lee como columnas fantasma
    'Unnamed: 4', 'Unnamed: 5', ... y el archivo era rechazado aunque los
    datos estuvieran bien. Si una columna 'Unnamed' SÍ tiene contenido, no se
    descarta: el validador lo reportará con el mensaje de formato.
    """
    fantasma = [
        col for col in df.columns
        if re.match(r"^unnamed:", str(col)) and _columna_vacia(df[col])
    ]
    return df.drop(columns=fantasma) if fantasma else df


def _es_fila_partida(row_id, concepto, unidad, cantidad) -> bool:
    """Determina si una fila es un encabezado de partida.

    Es partida si:
    - La columna id empieza con PARTIDA/SUBPARTIDA, O
    - id tiene texto pero concepto, unidad y cantidad están vacías.
    """
    if pd.isna(row_id):
        return False
    row_id_str = str(row_id).strip()
    if PATRON_PARTIDA.match(row_id_str):
        return True
    # Fila con id pero sin datos → también la tratamos como encabezado
    if row_id_str and pd.isna(concepto) and pd.isna(unidad) and pd.isna(cantidad):
        return True
    return False


def leer_entrada(ruta: str, partida_default: str = "GENERAL") -> list[dict]:
    """Lee el archivo de entrada y retorna una lista de conceptos.

    Detecta automáticamente:
    - Formato plano de 4 columnas (id | concepto | unidad | cantidad)
    - Formato de 5 columnas con partida explícita
    - Formato de 4 columnas con filas PARTIDA/SUBPARTIDA incrustadas

    Args:
        ruta: Ruta al archivo .xlsx de entrada.
        partida_default: Valor por defecto para 'partida' para conceptos
                         que aparecen antes de la primera partida.

    Returns:
        Lista de diccionarios con claves: id, partida, concepto, unidad, cantidad.

    Raises:
        FileNotFoundError: Si el archivo no existe.
        ValueError: Si las columnas no coinciden con el formato esperado.
    """
    ruta = Path(ruta)

    if not ruta.exists():
        raise FileNotFoundError(f"Archivo de entrada no encontrado: {ruta}")

    df = pd.read_excel(ruta)

    # Normalizar nombres de columna: strip, lowercase
    df.columns = [str(c).strip().lower() for c in df.columns]

    # Celdas con formato sobrante generan columnas 'Unnamed: N' vacías que
    # no alteran los datos: se descartan antes de validar el formato.
    df = _descartar_columnas_fantasma(df)
    columnas_encontradas = list(df.columns)

    # Detectar formato
    if columnas_encontradas == COLUMNAS_4:
        pass  # partida se resolverá por filas incrustadas o default
    elif columnas_encontradas == COLUMNAS_5:
        pass  # partida viene en su propia columna
    else:
        raise ValueError(
            f"Columnas inesperadas: {columnas_encontradas}\n"
            f"Se esperaba: {COLUMNAS_4} (4 cols) o {COLUMNAS_5} (5 cols)"
        )

    # Convertir a lista de diccionarios, detectando partidas
    conceptos = []
    partida_actual = partida_default
    subpartida_actual = None

    for _, row in df.iterrows():
        rid = row["id"]
        concepto = row.get("concepto")
        unidad = row.get("unidad")
        cantidad = row.get("cantidad")

        # ¿Es fila de partida/subpartida incrustada?
        if _es_fila_partida(rid, concepto, unidad, cantidad):
            titulo = str(rid).strip()
            if PATRON_PARTIDA.match(titulo) and "SUB" in titulo.upper()[:12]:
                subpartida_actual = titulo
            else:
                partida_actual = titulo
                subpartida_actual = None
            continue

        # Fila de concepto: validar campos obligatorios
        if pd.isna(rid) or pd.isna(concepto) or pd.isna(unidad) or pd.isna(cantidad):
            continue  # saltar filas totalmente vacías

        # Si hay columna 'partida' explícita (formato 5 col), usarla
        if "partida" in df.columns and pd.notna(row.get("partida")) and str(row.get("partida")).strip():
            partida_asignada = str(row["partida"]).strip()
        else:
            partida_asignada = partida_actual

        conceptos.append({
            "id": str(rid).strip(),
            "partida": partida_asignada,
            "subpartida": subpartida_actual,
            "concepto": str(concepto).strip(),
            "unidad": str(unidad).strip(),
            "cantidad": float(cantidad),
        })

    if not conceptos:
        raise ValueError(
            "No se encontraron conceptos válidos en el archivo. "
            "Verifica que existan filas con id, concepto, unidad y cantidad."
        )

    return conceptos


def validar_entrada(ruta: str) -> dict:
    """Valida el archivo de entrada sin cargarlo completamente.

    Retorna un dict con:
        - valido: bool
        - num_filas: int
        - num_partidas: int
        - columnas: list
        - error: str o None
    """
    try:
        df = pd.read_excel(ruta)
        df.columns = [str(c).strip().lower() for c in df.columns]
        conceptos = leer_entrada(ruta)
        partidas = sorted(set(c["partida"] for c in conceptos))
        return {
            "valido": True,
            "num_filas": len(conceptos),
            "num_partidas": len(partidas),
            "partidas": partidas,
            "columnas": list(conceptos[0].keys()) if conceptos else [],
            "error": None,
        }
    except Exception as e:
        return {
            "valido": False,
            "num_filas": 0,
            "num_partidas": 0,
            "partidas": [],
            "columnas": [],
            "error": str(e),
        }


# =============================================================================
# Generador de archivo de ejemplo (para testing)
# =============================================================================

def generar_entrada_ejemplo(ruta: str = "entrada_ejemplo.xlsx", num_conceptos: int = 10):
    """Genera un archivo XLSX de entrada de ejemplo para pruebas."""
    import random

    conceptos_ejemplo = [
        ("CONC-01", "PRELIMINARES", "Trazo y nivelación de terreno", "M2", random.uniform(100, 500)),
        ("CONC-02", "PRELIMINARES", "Limpieza y deshierbe de terreno", "M2", random.uniform(200, 800)),
        ("CONC-03", "CIMENTACION", "Excavación en cepa de 0 a 2m de profundidad", "M3", random.uniform(50, 200)),
        ("CONC-04", "CIMENTACION", "Plantilla de concreto f'c=100 kg/cm2 de 5cm espesor", "M2", random.uniform(100, 300)),
        ("CONC-05", "CIMENTACION", "Acero de refuerzo en cimentación fy=4200 kg/cm2", "KG", random.uniform(500, 2000)),
        ("CONC-06", "ESTRUCTURA", "Concreto en columnas f'c=250 kg/cm2", "M3", random.uniform(10, 50)),
        ("CONC-07", "ESTRUCTURA", "Cimbra acabado aparente en columnas", "M2", random.uniform(50, 200)),
        ("CONC-08", "ESTRUCTURA", "Acero de refuerzo en columnas fy=4200 kg/cm2", "KG", random.uniform(300, 1500)),
        ("CONC-09", "ALBAÑILERIA", "Muro de block hueco de concreto 15x20x40cm", "M2", random.uniform(100, 600)),
        ("CONC-10", "ALBAÑILERIA", "Aplanado fino en muros con mortero 1:4", "M2", random.uniform(200, 1000)),
    ]

    df = pd.DataFrame(conceptos_ejemplo, columns=["id", "partida", "concepto", "unidad", "cantidad"])
    df.to_excel(ruta, index=False)
    print(f"Archivo de ejemplo generado: {ruta} ({len(df)} conceptos)")
    return ruta


if __name__ == "__main__":
    # Test rápido
    generar_entrada_ejemplo("entrada_ejemplo.xlsx", 10)
    conceptos = leer_entrada("entrada_ejemplo.xlsx")
    print(f"\nLeídos {len(conceptos)} conceptos:")
    for c in conceptos:
        print(f"  {c['id']} | {c['partida']} | {c['concepto']} | {c['unidad']} | {c['cantidad']:.2f}")

    # Validación
    resultado = validar_entrada("entrada_ejemplo.xlsx")
    print(f"\nValidación: {resultado}")