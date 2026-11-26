"""Diagnóstico: mayores importes del presupuesto VENADO (con contrato)."""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[3]))  # raiz del repo
import sys, statistics
from src.costeo_client import ClienteCosteo
from src.indexador import Indexador
from src.lector_entrada import leer_entrada
from src.matcher import Matcher, normalizar_unidad

cli = ClienteCosteo()
print("Contrato:", cli.descripcion_contrato())
idx = Indexador(str(cli.indice_path)); idx.cargar()
conceptos = leer_entrada(r"H:\Mi unidad\ICEL\PUs\CASA VENADO\VENADO.xlsx")
res = Matcher(idx, umbral=55, piso=40, cliente=cli).emparejar(conceptos)

incl = [r for r in res if r["estado"] in ("confirmado","dudoso") and r.get("clave_sugerida")]
pu_uni = {}
for r in incl:
    pu = r.get("pu_sugerido") or 0
    if pu>0:
        pu_uni.setdefault(normalizar_unidad(r["unidad_entrada"]), []).append(pu)
med = {u:statistics.median(v) for u,v in pu_uni.items() if v}

def revisar(r):
    if r["estado"]=="confirmado": return False
    m = med.get(normalizar_unidad(r["unidad_entrada"]),0)
    return m>0 and (r.get("pu_sugerido") or 0) > 12*m

lineas=[]
for r in incl:
    pu = 0 if revisar(r) else (r.get("pu_sugerido") or 0)
    lineas.append((pu*r["cantidad"], r, revisar(r)))
lineas.sort(reverse=True, key=lambda x:x[0])
total = sum(l[0] for l in lineas)
print(f"Costo directo (con guardia): ${total:,.2f}\n")
print("TOP 20:")
print(f"{'IMPORTE':>16} {'ID':8} {'CANT':>9} {'UND_IN':6} {'UND_CAT':6} {'PU':>13} {'SC':>5} {'EST':5} CLAVE / DESC")
for imp, r, rev in lineas[:20]:
    flag = "REV" if rev else r["estado"][:4]
    print(f"{imp:>16,.0f} {r['id_entrada']:8} {r['cantidad']:>9.2f} {str(r['unidad_entrada']):6} {str(r.get('unidad_sugerida','')):6} {r.get('pu_sugerido',0):>13,.2f} {r['score']:>5.1f} {flag:5} {r['clave_sugerida'][:14]} / {r['descripcion_sugerida'][:40]}")
print("\nLineas REVISAR:", sum(1 for l in lineas if l[2]), "| Incluidas:", len(incl))