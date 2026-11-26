"""
apu_source.py
=============
Lee las MATRICES de precios unitarios del repo de costos (las que el catálogo
referencia con "VER EL DESGLOSE EN SU RESPECTIVA MATRIZ").

Ubicación (según rutas_costos.py del repo):
    <PONTEM_COSTOS_FUENTES>/<paquete>/MATRICES/<especialidad>/*.xlsx

Cada archivo tiene una hoja ('Sheet') con varios APUs apilados. Cada APU:
    - Clave, Unidad, Cantidad, Descripción, Precio unitario
    - Secciones: Materiales, Mano de Obra, Equipo, Herramienta, Auxiliares
    - Renglones: Clave | Descripción | Unidad | Cantidad | Costo unitario | Total

Uso:
    from src.apu_source import MatricesProvider
    prov = MatricesProvider()                 # descubre G:/Mi unidad/Costos/PU2026
    apu = prov.apu_de("10401-001", "OBRA CIVIL")
    # -> {"clave","descripcion","unidad","pu","insumos":[{tipo,clave,descripcion,unidad,cantidad,costo,total}]}
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Optional


def _norm(s) -> str:
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(s.lower().split())


SECCIONES = {
    "materiales": "material",
    "mano de obra": "mano_obra",
    "equipo": "equipo",
    "herramienta": "herramienta",
    "auxiliares": "auxiliar",
    "maquinaria": "maquinaria",
}

# Nombres de columna (normalizados) -> campo
COLS = {
    "clave": "clave",
    "descripcion": "descripcion",
    "unidad": "unidad",
    "cantidad": "cantidad",
    "costo unitario": "costo",
    "p unitario": "costo",
    "p. unitario": "costo",
    "precio unitario": "costo",
    "total": "total",
    "importe": "total",
}


def _num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        s = str(v).replace(",", "").replace("$", "").strip()
        return float(s)
    except Exception:
        return None


def _mapear_columnas(fila) -> dict:
    """Dada una fila (lista de valores), devuelve {campo: índice}."""
    mapa = {}
    for i, v in enumerate(fila):
        n = _norm(v)
        if n in COLS and COLS[n] not in mapa:
            mapa[COLS[n]] = i
    return mapa


def _es_fila_header(row) -> bool:
    txt = [_norm(x) for x in row if x not in (None, "")]
    tiene_desc = "descripcion" in txt
    tiene_cant = "cantidad" in txt
    tiene_costo = any(t in ("costo unitario", "p unitario", "p. unitario", "precio unitario") for t in txt)
    return tiene_desc and tiene_cant and tiene_costo


def _siguiente_valor(row, k):
    for kk in range(k + 1, len(row)):
        if row[kk] not in (None, ""):
            return row[kk]
    return None


def parsear_hoja(ws) -> list[dict]:
    """Extrae todos los APUs de una hoja. Devuelve lista de dicts."""
    filas = list(ws.iter_rows(values_only=True))
    n = len(filas)
    apus = []

    i = 0
    while i < n:
        row = filas[i]
        if row is None:
            i += 1
            continue
        a = _norm(row[0]) if row else ""
        # Inicio de bloque: título "Análisis de precios" o una fila de metadatos "Clave:"
        es_inicio = a.startswith("analisis de precio") or (a.rstrip(":") == "clave")
        # La fila de header NO es inicio (contiene Descripción+Cantidad+Costo)
        if _es_fila_header(row):
            es_inicio = False

        if not es_inicio:
            i += 1
            continue

        apu = {"clave": None, "descripcion": "", "unidad": "", "pu": None, "insumos": []}
        j = i
        header_idx = None
        while j < n and j < i + 30:
            rowj = filas[j]
            if rowj is None:
                j += 1
                continue
            if _es_fila_header(rowj):
                header_idx = j
                break
            # Metadatos (etiquetas con ':'): clave, unidad, precio unitario
            for k, v in enumerate(rowj):
                nv = _norm(v)
                if nv.startswith("clave"):
                    val = _siguiente_valor(rowj, k)
                    if val is not None and _norm(val) != "descripcion":
                        apu["clave"] = str(val).strip()
                elif nv.startswith("unidad"):
                    val = _siguiente_valor(rowj, k)
                    if val is not None and _norm(val) != "cantidad":
                        apu["unidad"] = str(val).strip()
                elif nv.startswith("precio unitario"):
                    val = _siguiente_valor(rowj, k)
                    num = _num(val)
                    if num is not None:
                        apu["pu"] = num
            # Descripción: primer texto largo en col A que no sea título/etiqueta
            if (not apu["descripcion"] and j > i and rowj and rowj[0] not in (None, "")
                    and not _norm(rowj[0]).startswith(("analisis", "clave", "unidad", "cantidad", "precio", "total"))
                    and len(str(rowj[0]).strip()) > 8):
                apu["descripcion"] = str(rowj[0]).strip()
            j += 1

        if header_idx is None:
            i += 1
            continue

        cols = _mapear_columnas(filas[header_idx])
        c_clave = cols.get("clave")
        c_desc = cols.get("descripcion")
        c_uni = cols.get("unidad")
        c_cant = cols.get("cantidad")
        c_costo = cols.get("costo")
        c_total = cols.get("total")

        seccion_actual = None
        k = header_idx + 1
        while k < n:
            rowk = filas[k]
            if rowk is None:
                k += 1
                continue
            an = _norm(rowk[0]) if rowk else ""
            # fin de bloque: nueva clave, o "costo directo" (en cualquier columna)
            if an.rstrip(":") == "clave":
                break
            if any(_norm(x) == "costo directo" for x in rowk if x is not None):
                break
            if an.startswith("analisis de precio"):
                break
            if an in SECCIONES:
                seccion_actual = SECCIONES[an]
                k += 1
                continue
            if an.startswith("total de"):
                k += 1
                continue
            if c_clave is not None and c_desc is not None:
                clave_i = rowk[c_clave] if c_clave < len(rowk) else None
                desc_i = rowk[c_desc] if c_desc < len(rowk) else None
                cant_i = _num(rowk[c_cant]) if (c_cant is not None and c_cant < len(rowk)) else None
                costo_i = _num(rowk[c_costo]) if (c_costo is not None and c_costo < len(rowk)) else None
                total_i = _num(rowk[c_total]) if (c_total is not None and c_total < len(rowk)) else None
                if (clave_i not in (None, "") and desc_i not in (None, "")
                        and cant_i is not None and costo_i is not None
                        and seccion_actual is not None):
                    apu["insumos"].append({
                        "tipo": seccion_actual,
                        "clave": str(clave_i).strip(),
                        "descripcion": str(desc_i).strip(),
                        "unidad": str(rowk[c_uni]).strip() if (c_uni is not None and c_uni < len(rowk) and rowk[c_uni] is not None) else "",
                        "cantidad": cant_i,
                        "costo": costo_i,
                        "total": total_i if total_i is not None else round(cant_i * costo_i, 4),
                    })
            k += 1

        if apu["clave"] and apu["clave"].upper() != "DESCRIPCION":
            apus.append(apu)
        i = k if k > i else i + 1

    return apus


def parsear_archivo(path: str) -> list[dict]:
    import openpyxl
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return []
    out = []
    try:
        for name in wb.sheetnames:
            try:
                out.extend(parsear_hoja(wb[name]))
            except Exception:
                continue
    finally:
        wb.close()
    return out


# ---------------------------------------------------------------------------
# Descubrimiento de FUENTES (PU2026)
# ---------------------------------------------------------------------------

def descubrir_fuentes(repo: Optional[str] = None) -> Optional[Path]:
    """Localiza la carpeta de matrices (PU2026)."""
    # 1) variable de entorno
    env = os.environ.get("PONTEM_COSTOS_FUENTES")
    if env and Path(env).exists():
        return Path(env)
    # 2) rutas_costos.py del repo
    if repo:
        rc = Path(repo) / "costos_unitarios" / "rutas_costos.py"
        if rc.exists():
            txt = rc.read_text(encoding="utf-8", errors="ignore")
            m = re.search(r'FUENTES\s*=\s*Path\(os\.environ\.get\([^,]+,\s*(DRIVE\s*/\s*"[^"]+")', txt)
            m2 = re.search(r'PONTEM_COSTOS_FUENTES"\s*,\s*DRIVE\s*/\s*"([^"]+)"', txt)
            m3 = re.search(r'DRIVE\s*=\s*Path\(os\.environ\.get\([^,]+,\s*r"([^"]+)"', txt)
            if m3:
                drive = Path(m3.group(1))
                if m2:
                    cand = drive / m2.group(1)
                    if cand.exists():
                        return cand
    # 3) defaults conocidos
    for cand in [Path(r"G:/Mi unidad/Costos/PU2026"),
                 Path(r"C:/Pontem_dev/PROYECTOS_PYTHON/Proy_Codigos/Costos/PU2026")]:
        if cand.exists():
            return cand
    return None


def listar_matrices(fuentes: Path) -> list[Path]:
    """Lista todos los .xlsx bajo */MATRICES/**."""
    out = []
    for p in fuentes.glob("*/MATRICES/**/*.xlsx"):
        if p.name.startswith("~$"):
            continue
        out.append(p)
    return out


class MatricesProvider:
    """Índice perezoso clave -> APU, con caché en disco."""

    def __init__(self, repo: Optional[str] = None, cache_dir: Optional[str] = None,
                 verbose: bool = False):
        self.repo = repo
        self.fuentes = descubrir_fuentes(repo)
        self.verbose = verbose
        base = Path(__file__).resolve().parent.parent
        self.cache_dir = Path(cache_dir) if cache_dir else base / "cache"
        self._index: dict[str, dict] = {}
        self._cargado = False

    @property
    def disponible(self) -> bool:
        return self.fuentes is not None and self.fuentes.exists()

    def _cache_path(self) -> Path:
        return self.cache_dir / "apu_index.json"

    def _meta_path(self) -> Path:
        return self.cache_dir / "apu_index.meta.json"

    def _fingerprint_fuentes(self) -> str:
        """Huella de las matrices fuente.

        Hash de (ruta relativa, mtime_ns, tamaño) de cada archivo: detecta
        cualquier cambio (editar, agregar o borrar una matriz). Un simple
        "mtime máximo" no serviría, porque editar una matriz que no sea la
        más reciente dejaría el cache vencido sin detectarlo.
        """
        if not self.fuentes or not self.fuentes.exists():
            return "sin-fuentes"
        archivos = listar_matrices(self.fuentes)
        if not archivos:
            return "sin-archivos"
        h = hashlib.sha1()
        for p in sorted(archivos):
            try:
                st = p.stat()
                rel = p.relative_to(self.fuentes)
            except (OSError, ValueError):
                continue
            h.update(f"{rel}|{st.st_mtime_ns}|{st.st_size}".encode("utf-8", "replace"))
        return f"{len(archivos)}:{h.hexdigest()[:16]}"

    def _cache_al_dia(self) -> bool:
        """El cache existe y corresponde a las matrices actuales."""
        cp = self._cache_path()
        if not cp.exists():
            return False
        mp = self._meta_path()
        if not mp.exists():
            # Cache heredado (sin meta): se regenera para no servir datos viejos
            return False
        try:
            meta = json.loads(mp.read_text(encoding="utf-8"))
        except Exception:
            return False
        return meta.get("fingerprint") == self._fingerprint_fuentes()

    def _guardar_cache(self, index: dict):
        try:
            self.cache_dir.mkdir(parents=True, exist_ok=True)
            self._cache_path().write_text(
                json.dumps(index, ensure_ascii=False), encoding="utf-8")
            self._meta_path().write_text(
                json.dumps({
                    "generado": datetime.now().isoformat(timespec="seconds"),
                    "fingerprint": self._fingerprint_fuentes(),
                    "apus": len(index),
                }, ensure_ascii=False),
                encoding="utf-8")
        except Exception:
            pass

    def cargar_indice(self, forzar: bool = False) -> int:
        """Construye/carga el índice completo clave->APU.

        El cache es LOCAL y no se versiona (ver .gitignore): se regenera solo
        cuando las matrices fuente cambian.
        """
        if self._cargado and not forzar:
            return len(self._index)
        if not forzar and self._cache_al_dia():
            try:
                self._index = json.loads(self._cache_path().read_text(encoding="utf-8"))
                self._cargado = True
                return len(self._index)
            except Exception:
                pass
        if not self.disponible:
            self._cargado = True
            return 0

        archivos = listar_matrices(self.fuentes)
        if self.verbose:
            print(f"  Matrices encontradas: {len(archivos)} archivos")
        index: dict[str, dict] = {}
        for i, path in enumerate(archivos, 1):
            if self.verbose and i % 50 == 0:
                print(f"    ... {i}/{len(archivos)}")
            for apu in parsear_archivo(str(path)):
                index[apu["clave"]] = apu
        self._index = index
        self._cargado = True
        self._guardar_cache(index)
        return len(index)

    def apu_de(self, clave: str, paquete: Optional[str] = None) -> Optional[dict]:
        if not self._cargado:
            self.cargar_indice()
        return self._index.get(str(clave))

    def stats(self) -> dict:
        return {"fuentes": str(self.fuentes) if self.fuentes else None,
                "apus": len(self._index),
                "disponible": self.disponible}


if __name__ == "__main__":
    prov = MatricesProvider(verbose=True)
    print("Fuentes:", prov.fuentes)
    if prov.disponible:
        n = prov.cargar_indice()
        print("APUs indexados:", n)
        for clave in list(prov._index.keys())[:3]:
            a = prov._index[clave]
            print(f"  {clave} | {a['unidad']} | ${a['pu']} | {a['descripcion'][:50]} | insumos={len(a['insumos'])}")
            for ins in a["insumos"][:3]:
                print("     ", ins["tipo"], ins["clave"], ins["descripcion"][:30], ins["cantidad"], ins["costo"])