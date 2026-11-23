"""
matcher.py
==========
Emparejador de conceptos usando el contrato oficial del repo de costos (Nivel 3).

Estrategia:
1. Usa ClienteCosteo.buscar() (buscador oficial Nivel 3) como fuente primaria de candidatos.
2. Re-rankea con IDF + token_set_ratio + cobertura ponderada.
3. Filtro duro por unidad incompatible.

Uso:
    from src.matcher import Matcher
    from src.costeo_client import ClienteCosteo

    cliente = ClienteCosteo()  # descubre repo automáticamente
    m = Matcher(cliente=cliente, umbral=55, piso=40)
    resultados = m.emparejar(conceptos_entrada)
"""

import re
import math
import unicodedata
from collections import defaultdict, Counter
from typing import Optional, List, Dict

from rapidfuzz import fuzz, process

from src.costeo_client import ClienteCosteo


# =============================================================================
# Normalización
# =============================================================================

def normalizar(texto: str) -> str:
    """Normaliza texto: sin acentos, minúsculas, sin puntuación, espacios simples."""
    texto = str(texto or "").strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = texto.encode("ascii", "ignore").decode("ascii").lower()
    texto = re.sub(r"[^a-z0-9 ]", " ", texto)
    texto = re.sub(r"\s+", " ", texto)
    return texto


def normalizar_unidad(unidad: str) -> str:
    """Normaliza unidad (conserva mayúsculas comunes: M2, M3, ML, KG, PZA, etc.)."""
    unidad = unidad.strip().upper()
    sinonimos = {
        "METRO CUADRADO": "M2", "METROS CUADRADOS": "M2", "M2": "M2",
        "METRO CUBICO": "M3", "METROS CUBICOS": "M3", "M3": "M3",
        "METRO LINEAL": "ML", "METROS LINEALES": "ML", "ML": "ML",
        "KILOGRAMO": "KG", "KILOGRAMOS": "KG", "KG": "KG",
        "PIEZA": "PZA", "PIEZAS": "PZA", "PZA": "PZA",
        "LOTE": "LOTE", "SISTEMA": "SIS", "SIS": "SIS",
        "JORNADA": "JOR", "JOR": "JOR", "SACO": "SAC", "SACOS": "SACO",
        "TONELADA": "TON", "TON": "TON", "LITRO": "LT", "LT": "LT",
    }
    return sinonimos.get(unidad, unidad)


# Stopwords en español + palabras de relleno típicas de presupuestos
STOPWORDS = set("""
de del la el los las en para con por y o u a al un una uno unos unas
tipo segun proyecto obra obras m2 m3 ml kg pza lote sistema jor jornal
suministro suministrar habilitado habilitar instalacion instalar colocacion colocar
elaboracion elaborar aplicacion aplicar acabado acabados material materiales
mano equipo herramienta necesario necesarios correcta correcto ejecucion ejecutar
pieza piezas servicio servicios incluye incluyendo incluidas incluidos hasta
mediante sobre entre cada todo toda todos todas su sus del al como mas
cm mm mts mts2 mts3 m m2 m3 kg pza pzas lote lt saco sacos ton tonelada
suministro suministrar habilitado habilitar instalacion instalar colocacion colocar
elaboracion elaborar aplicacion aplicar acabado acabados material materiales
mano equipo herramienta necesario necesarios correcta correcto ejecucion ejecutar
pieza piezas servicio servicios incluye incluyendo incluidas incluidos hasta
mediante sobre entre cada todo toda todos todas su sus del al como mas
cm mm mts mts2 mts3 m m2 m3 kg pza pzas lote lt saco sacos ton tonelada
suministro suministrar habilitado habilitar instalacion instalar colocacion colocar
elaboracion elaborar aplicacion aplicar acabado acabados material materiales
mano equipo herramienta necesario necesarios correcta correcto ejecucion ejecutar
pieza piezas servicio servicios incluye incluyendo incluidas incluidos hasta
mediante sobre entre cada todo toda todos todas su sus del al como mas
cm mm mts mts2 mts3 m m2 m3 kg pza pzas lote lt saco sacos ton tonelada
""".split())


def tokens(texto: str) -> list[str]:
    """Tokens significativos (sin stopwords, longitud > 1)."""
    return [t for t in normalizar(texto).split() if t not in STOPWORDS and len(t) > 1]


# =============================================================================
# Matcher (usa ClienteCosteo como fuente)
# =============================================================================

class Matcher:
    """Emparejador mejorado usando ClienteCosteo (buscador oficial Nivel 3)."""

    def __init__(self, cliente: "ClienteCosteo", umbral: int = 55, piso: int = 40):
        """
        Args:
            cliente: ClienteCosteo ya inicializado (con repo descubierto).
            umbral: score mínimo (0-100) para estado 'confirmado'.
            piso: score mínimo (0-100) para incluir como 'dudoso' (debajo -> sin_match).
        """
        self.cliente = cliente
        self.umbral = max(0, min(100, umbral))
        self.piso = max(0, min(100, piso))

        # Pre-cargar todos los conceptos del catálogo para scoring IDF
        self._conceptos = cliente.todos()
        self._tstr: list[str] = []
        self._tset: list[set] = []
        self._unidad: list[str] = []
        self._por_unidad: dict[str, list[int]] = defaultdict(list)
        self._por_desc = {}
        self._clave_a_idx: dict[str, int] = {}

        for i, c in enumerate(self._conceptos):
            tk = tokens(c.get("descripcion"))
            tset = set(tk)
            tstr = " ".join(sorted(tset))
            uni = normalizar_unidad(c.get("unidad"))
            self._tstr.append(tstr)
            self._tset.append(tset)
            self._unidad.append(uni)
            self._por_unidad[uni].append(i)
            dn = normalizar(c.get("descripcion"))
            if dn:
                self._por_desc[dn] = c

        # IDF: rareza de cada token en el catálogo
        df = Counter()
        for tset in self._tset:
            for t in tset:
                df[t] += 1
        n_total = max(1, len(self._conceptos))
        self._idf = {t: math.log(n_total / (1 + d)) for t, d in df.items()}
        idf_vals = sorted(self._idf.values())
        self._idf_mediana = idf_vals[len(idf_vals) // 2] if idf_vals else 0.0

    def emparejar(self, conceptos_entrada: list[dict]) -> list[dict]:
        """Empareja todos los conceptos de entrada. Retorna lista de resultados."""
        return [self._emparejar_uno(c) for c in conceptos_entrada]

    def _emparejar_uno(self, ce: dict) -> dict:
        id_entrada = ce.get("id", "")
        partida = ce.get("partida", "")
        subpartida = ce.get("subpartida", "")
        desc_entrada = ce.get("concepto", "")
        unidad_entrada = ce.get("unidad", "")
        cantidad = ce.get("cantidad", 0)

        base = {
            "id_entrada": id_entrada,
            "partida": partida,
            "subpartida": subpartida,
            "concepto_entrada": desc_entrada,
            "unidad_entrada": unidad_entrada,
            "cantidad": cantidad,
        }

        desc_norm = normalizar(desc_entrada)
        q_uni = normalizar_unidad(unidad_entrada)
        q_tokens = tokens(desc_entrada)
        q_set = set(q_tokens)
        q_str = " ".join(sorted(q_set))

        base = {
            "id_entrada": id_entrada,
            "partida": partida,
            "subpartida": subpartida,
            "concepto_entrada": desc_entrada,
            "unidad_entrada": unidad_entrada,
            "cantidad": cantidad,
        }

        # 1. Match exacto por descripción normalizada
        if desc_norm in self._por_desc:
            c = self._por_desc[desc_norm]
            return {
                **base,
                "clave_sugerida": c.get("clave", ""),
                "descripcion_sugerida": c.get("descripcion", ""),
                "unidad_sugerida": c.get("unidad", ""),
                "score": 100.0,
                "metodo": "exacto",
                "estado": "confirmado",
                "pu_sugerido": c.get("precio_unitario", 0),
                "candidatos": [],
            }

        # 2. Buscar candidatos con ClienteCosteo.buscar() (fuente oficial Nivel 3)
        candidatos_oficiales = []
        if self.cliente:
            try:
                oficiales = self.cliente.buscar(desc_entrada, limite=10)
                for o in oficiales:
                    idx = self._clave_a_idx.get(o.get("clave"))
                    if idx is not None:
                        candidatos_oficiales.append(idx)
            except Exception:
                pass

        # 2b. Búsqueda fuzzy con rapidfuzz (fallback)
        if not self._tstr:
            return {**base, "clave_sugerida": "", "descripcion_sugerida": "",
                    "unidad_sugerida": "", "score": 0, "metodo": "sin_match",
                    "estado": "sin_match", "pu_sugerido": 0, "candidatos": []}

        q_str = " ".join(sorted(set(tokens(desc_entrada))))
        resultados_fuzzy = process.extract(
            q_str, self._tstr, scorer=fuzz.token_set_ratio, limit=10
        )

        # Combinar candidatos oficiales + fuzzy
        candidatos_idx = set()
        for idx in candidatos_oficiales:
            candidatos_idx.add(idx)
        for _, _, idx in resultados_fuzzy:
            candidatos_idx.add(idx)

        # Evaluar cada candidato
        candidatos = []
        for idx in candidatos_idx:
            c = self._conceptos[idx]
            # Filtro duro por unidad
            if self._unidad[idx] != normalizar_unidad(unidad_entrada):
                continue

            # Score combinado: token_set_ratio + cobertura ponderada por IDF
            sc = fuzz.token_set_ratio(normalizar(c.get("descripcion")), desc_norm)
            hay = sum(1 for t in q_set if t in self._tset[idx] or any(t in x or x in t for x in self._tset[idx]))
            cov = hay / len(q_set) if q_set else 0
            comb = 0.6 * sc + 0.4 * cov * 100
            bonus = 8 if any(self._idf.get(t, 0) >= self._idf_mediana for t in q_set & self._tset[idx]) else 0
            comb = min(100, comb + bonus)
            candidatos.append((comb, sc, idx))

        if not candidatos:
            return {**base, "clave_sugerida": "", "descripcion_sugerida": "",
                    "score": 0, "metodo": "sin_match", "estado": "sin_match",
                    "pu_sugerido": 0, "candidatos": []}

        candidatos.sort(reverse=True, key=lambda x: x[0])
        mejor_score, _, idx_best = candidatos[0]
        best = self._conceptos[idx_best]

        if mejor_score >= self.umbral:
            estado = "confirmado"
        elif mejor_score >= self.piso:
            estado = "dudoso"
        else:
            estado = "sin_match"

        # Candidatos top-5 para revisión
        top_candidatos = []
        for comb, sc, idx in candidatos[:5]:
            c = self._conceptos[idx]
            top_candidatos.append({
                "clave": c.get("clave", ""),
                "descripcion": c.get("descripcion", ""),
                "unidad": c.get("unidad", ""),
                "precio_unitario": c.get("precio_unitario", 0),
                "score": round(comb, 1),
            })

        return {
            **base,
            "clave_sugerida": self._conceptos[idx_best].get("clave", ""),
            "descripcion_sugerida": self._conceptos[idx_best].get("descripcion", ""),
            "unidad_sugerida": self._conceptos[idx_best].get("unidad", ""),
            "score": round(mejor_score, 1),
            "metodo": "oficial" if idx_best in candidatos_oficiales else "fuzzy",
            "estado": estado,
            "pu_sugerido": self._conceptos[idx_best].get("precio_unitario", 0),
            "candidatos": top_candidatos,
        }

    def stats_matcheo(self, resultados: list[dict]) -> dict:
        total = len(resultados)
        exactos = sum(1 for r in resultados if r["metodo"] == "exacto")
        fuzzy = sum(1 for r in resultados if r["metodo"] == "fuzzy")
        sin_match = sum(1 for r in resultados if r["metodo"] == "sin_match")
        confirmados = sum(1 for r in resultados if r["estado"] == "confirmado")
        dudosos = sum(1 for r in resultados if r["estado"] == "dudoso")
        return {
            "total": total,
            "exactos": exactos,
            "fuzzy": fuzzy,
            "sin_match": sin_match,
            "confirmados": confirmados,
            "dudosos": dudosos,
            "pct_confirmados": round(confirmados / total * 100, 1) if total else 0,
        }


# =============================================================================
# Test rápido
# =============================================================================

if __name__ == "__main__":
    import sys
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

    from src.costeo_client import ClienteCosteo

    idx = Indexador(r"C:\Pontem_dev\costeo\costos.db")
    idx.cargar()
    print(f"Catalogo: {idx.total} conceptos")

    from src.costeo_client import ClienteCosteo
    cli = ClienteCosteo()
    m = Matcher(cliente=cli, umbral=55)
    resultados = m.emparejar([
        {"id": "CIM-01", "concepto": "Excavacion en cepa material B", "unidad": "M3", "cantidad": 80},
        {"id": "CIM-02", "concepto": "Muro de block 15cm", "unidad": "M2", "cantidad": 500},
    ])
    for r in resultados:
        print(f"{r['id_entrada']:8s} | {r['estado']:11s} | {r['score']:5.1f} | {r['clave_sugerida']} | {r['descripcion_sugerida'][:50]}")