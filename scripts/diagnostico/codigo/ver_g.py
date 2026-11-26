
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[3]))  # raiz del repo
import sys
from src.lector_entrada import leer_entrada
cs = leer_entrada(r"H:\Mi unidad\ICEL\PUs\CASA VENADO\VENADO.xlsx")
for c in cs:
    if c["id"].startswith("G-"):
        print(f"{c['id']:6} | {c['unidad']:5} | {c['cantidad']:>7} | {c['concepto'][:95]}")