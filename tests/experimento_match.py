"""Experimento: matcher mejorado con filtro de unidad + stopwords."""
import sys, re
sys.path.insert(0, ".")
sys.path.insert(0, r"C:\Pontem_dev\costeo")

import unicodedata
from rapidfuzz import fuzz

CAT = r"C:\Pontem_dev\costeo\costos_unitarios\catalogo_costos.json"

def norm(s):
    s = str(s or "")
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii").lower()
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    return " ".join(s.split())

STOP = set("""de del la el los las en para con por y o a al un una uno tipo incluye incluyendo
segun segun proyecto obras obra cm mm m mts m2 m3 kg pza ml lote sistema jornada jor sac ton lt
suministro habilitado instalacion colocacion hecho obra elaboracion aplicacion acabado
material materiales mano equipo herramienta necesario necesarios correcta ejecucion
pieza piezas servicio servicios""".split())

def tokens(s):
    return [t for t in norm(s).split() if t not in STOP and len(t) > 1]

import json
cs = json.load(open(CAT, encoding="utf-8"))["conceptos"]

def unidad_norm(u):
    u = norm(u)
    sin = {"m2":"m2","m3":"m3","ml":"ml","kg":"kg","pza":"pza","lote":"lote","sistema":"sistema","jor":"jor","ton":"ton","lt":"lt","saco":"saco","m":"m"}
    return sin.get(u, u)

from src.lector_entrada import leer_entrada
conceptos = leer_entrada(r"H:\Mi unidad\ICEL\PUs\CASA VENADO\VENADO.xlsx")
muestras = [conceptos[i] for i in [0, 2, 4, 6, 11, 12]]

for m in muestras:
    q = tokens(m["concepto"])
    uni = unidad_norm(m["unidad"])
    print("=" * 90)
    print(f"ENTRADA: {m['concepto'][:75]}  [{uni}]")
    cands = []
    for c in cs:
        if unidad_norm(c.get("unidad")) != uni and uni not in ("", "m"):
            continue
        ct = tokens(c.get("descripcion"))
        if not ct:
            continue
        # token_set_ratio sobre tokens filtrados
        qs = " ".join(sorted(set(q)))
        cstr = " ".join(sorted(set(ct)))
        sc = fuzz.token_set_ratio(qs, cstr)
        # cobertura: fraccion de tokens de query presentes en desc
        hay = sum(1 for t in q if any(t in x or x in t for x in ct))
        cov = hay / max(1, len(q))
        cands.append((sc * 0.6 + cov * 100 * 0.4, sc, cov, c))
    cands.sort(reverse=True, key=lambda x: x[0])
    for sc_comb, sc, cov, c in cands[:3]:
        print(f"  comb={sc_comb:5.1f} tset={sc:5.1f} cov={cov:.2f} | {c['clave']} | {c['unidad']} | ${c['precio_unitario']} | {c['descripcion'][:55]}")
    print()
