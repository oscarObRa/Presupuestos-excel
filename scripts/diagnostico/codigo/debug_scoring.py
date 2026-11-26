"""Debug de scoring para ENT-05 y ENT-09."""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[3]))  # raiz del repo
import sys
from src.indexador import Indexador
from src.lector_entrada import leer_entrada
from src.matcher import Matcher, tokens, normalizar, normalizar_unidad
from rapidfuzz import fuzz

idx = Indexador(r"C:\Pontem_dev\costeo\costos_unitarios\catalogo_costos.json")
idx.cargar()
conceptos = leer_entrada(r"H:\Mi unidad\ICEL\PUs\CASA VENADO\VENADO.xlsx")

for ce in conceptos:
    if ce["id"] not in ("ENT-05", "ENT-09"):
        continue
    print("=" * 80)
    print(ce["id"], "::", ce["concepto"][:70], "[", ce["unidad"], "]")
    q = tokens(ce["concepto"]); qs = set(q)
    print("  q_tokens:", sorted(qs))
    # buscar el concepto que salio elegido
    clave_obj = {"ENT-05": "5000066", "ENT-09": "10501-515"}[ce["id"]]
    c = idx.por_clave(clave_obj)
    if c:
        ct = tokens(c["descripcion"]); cs = set(ct); tc = set(ct)
        inter = qs & cs
        sc = fuzz.token_set_ratio(" ".join(sorted(qs)), " ".join(sorted(tc)))
        hay = sum(1 for t in qs if t in tc or any(t in x or x in t for x in tc))
        cov = hay / len(qs)
        print("  elegido:", clave_obj, "::", c["descripcion"][:60])
        print("  c_tokens:", sorted(cs))
        print(f"  interseccion={inter} tset={sc} hay={hay}/{len(qs)} cov={cov:.2f} comb={0.6*sc+0.4*cov*100:.1f}")
