"""Analiza la estructura del archivo VENADO.xlsx"""
import pandas as pd

df = pd.read_excel(r"H:\Mi unidad\ICEL\PUs\CASA VENADO\VENADO.xlsx")
print(f"Columnas: {list(df.columns)}")
print(f"Filas totales: {len(df)}")

# Detectar partidas y subpartidas
partidas = df[df["ID"].str.startswith("PARTIDA", na=False) | df["ID"].str.startswith("SUB", na=False)]
print(f"\n=== PARTIDAS Y SUBPARTIDAS ({len(partidas)}) ===")
for idx, row in partidas.iterrows():
    es_partida = str(row["ID"]).startswith("PARTIDA")
    tipo = "PARTIDA" if es_partida else "SUBPARTIDA"
    print(f"  Fila {idx:3d} [{tipo}]: {row['ID']}")

# Contar conceptos por partida
print("\n=== CONCEPTOS POR PARTIDA ===")
partida_actual = "SIN PARTIDA"
contador = 0
for idx, row in df.iterrows():
    rid = str(row["ID"])
    if rid.startswith("PARTIDA") or rid.startswith("SUB"):
        if contador > 0:
            print(f"  {partida_actual[:60]}: {contador} conceptos")
        partida_actual = rid
        contador = 0
    elif pd.notna(row.get("Concepto")):
        contador += 1
if contador > 0:
    print(f"  {partida_actual[:60]}: {contador} conceptos")

# Ver IDs únicos no-partida
ids_conceptos = df[~df["ID"].str.startswith("PARTIDA", na=False) & ~df["ID"].str.startswith("SUB", na=False)]["ID"].dropna()
print(f"\nTotal conceptos reales: {len(ids_conceptos)}")
print(f"IDs de concepto (primeros 10): {list(ids_conceptos.head(10))}")