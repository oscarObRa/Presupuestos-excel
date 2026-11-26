"""Top importes del budget con guardia aplicada."""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[3]))  # raiz del repo
import sys, statistics
from src.indexador import Indexador
from src.lector_entrada import leer_entrada
from src.matcher import Matcher, normalizar_unidad

idx = Indexador(r"C:\Pontem_dev\costeo\costos_unitarios\catalogo_costos.json")
idx.cargar()
conceptos = leer_entrada(r"H:\Mi unidad\ICEL\PUs\CASA VENADO\VENADO.xlsx")
res = Matcher(idx, umbral=55, piso=40).emparejar(conceptos)

incl = [r for r in res if r["estado"] in ("confirmado","dudoso") and r.get("clave_sugerida")]
pu_uni = {}
for r in incl:
    pu = r.get("pu_sugerido") or 0
    if pu>0:
        pu_uni.setdefault(normalizar_unidad(r["unidad_entrada"]), []).append(pu)
med = {u:statistics.median(v) for u,v in pu_uni.items() if v}

def revisar(r):
    if r["estado"]=="confirmado": return False
    med_u = med.get(normalizar_unidad(r["unidad_entrada"]),0)
    return med_u>0 and (r.get("pu_sugerido") or 0) > 12*med_u

lineas=[]
for r in incl:
    pu = 0 if revisar(r) else (r.get("pu_sugerido") or 0)
    imp = pu * r["cantidad"]
    lineas.append((imp, r, revisar(r)))

lineas.sort(reverse=True, key=lambda x:x[0])
total = sum(l[0] for l in lineas)
print(f"Total costo directo (con guardia): ${total:,.2f}")
print("\nTop 15 importes:")
for imp, r, rev in lineas[:15]:
    flag = "REVISAR" if rev else r["estado"][:4]
    print(f"  ${imp:>14,.2f} | {flag:7s} | score={r['score']:5.1f} | {r['id_entrada']:8s} | {r['clave_sugerida']:12s} | {r['descripcion_sugerida'][:42]}")
print(f"\nLineas marcadas REVISAR: {sum(1 for l in lineas if l[2])}")
print(f"Lineas incluidas: {len(incl)}")
