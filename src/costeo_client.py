"""
costeo_client.py
================
Cliente que consume el sistema de costos PONTEM **según el contrato** documentado
en docs/consumo-costos.md.

- Descubre el repo de costos entre carpetas hermanas (o ancestros), buscando:
    <repo>/buscador_costos.py   (Nivel 3 — buscador portátil)
    <repo>/costos_unitarios/catalogo_costos.json   (Nivel 1/3 — índice)
- Si existe `buscador_costos.py`, lo importa y usa su API oficial `buscar()`
  (Nivel 3). Si no, cae a lectura directa del JSON (Nivel 1).
- Expone `buscar(texto, limite)` y `por_clave(clave)`.

Uso:
    from src.costeo_client import ClienteCosteo

    cliente = ClienteCosteo()          # descubre el repo solo
    resultados = cliente.buscar("muro de block", limite=10)

    # Nivel 1 (sin buscador): misma API, lectura directa del JSON
"""

from __future__ import annotations

import importlib.util
import json
import re
import sys
import unicodedata
from pathlib import Path
from typing import Optional


# ---------------------------------------------------------------------------
# Normalización local (para el fallback Nivel 1)
# ---------------------------------------------------------------------------

def _norm(s) -> str:
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(s.lower().split())


# ---------------------------------------------------------------------------
# Descubrimiento del repo
# ---------------------------------------------------------------------------

def _puntaje_repo(d: Path) -> int:
    """Puntúa qué tan probable es que `d` sea el repo de costos real.

    Prioriza tener el buscador oficial; penaliza un JSON suelto aislado.
    """
    try:
        if not d.is_dir():
            return 0
    except Exception:
        return 0
    score = 0
    if (d / "buscador_costos.py").exists():
        score += 100
    if (d / "costos_unitarios" / "catalogo_costos.json").exists():
        score += 50
    elif (d / "catalogo_costos.json").exists():
        score += 5  # JSON suelto: baja prioridad
    return score


def _raiz_app() -> Path:
    """Carpeta raíz de esta aplicación (para excluirla de la búsqueda)."""
    return Path(__file__).resolve().parents[1]


def _dentro_de(hijo: Path, padre: Path) -> bool:
    try:
        hijo.resolve().relative_to(padre.resolve())
        return True
    except Exception:
        return False


def descubrir_repo_costeo(inicio: Optional[str | Path] = None,
                          max_niveles: int = 6) -> Optional[Path]:
    """Busca el repo de costos entre carpetas hermanas y ancestros.

    Prioriza el que tenga `buscador_costos.py` (contrato Nivel 3). Excluye la
    carpeta de esta misma aplicación y sus subcarpetas.
    """
    base = Path(inicio).resolve() if inicio else Path(__file__).resolve().parent
    raiz_app = _raiz_app()

    candidatos: dict[Path, int] = {}
    actual = base
    for _ in range(max_niveles):
        # la propia carpeta
        sc = _puntaje_repo(actual)
        if sc > 0:
            candidatos[actual] = max(candidatos.get(actual, 0), sc)
        # hermanas
        try:
            padre = actual.parent
            if padre.exists():
                for d in padre.iterdir():
                    sc = _puntaje_repo(d)
                    if sc > 0:
                        candidatos[d] = max(candidatos.get(d, 0), sc)
        except Exception:
            pass

        if actual.parent == actual:
            break
        actual = actual.parent

    # Excluir la app y su contenido
    candidatos = {
        d: s for d, s in candidatos.items()
        if not _dentro_de(d, raiz_app) and d != raiz_app
    }
    if not candidatos:
        return None

    # Mejor puntaje; a igual puntaje, el más cercano a `base`
    mejor = max(candidatos.items(), key=lambda kv: (kv[1], -len(str(kv[0]))))
    return mejor[0]


def ruta_indice(repo: Path) -> Optional[Path]:
    """Devuelve la ruta al catalogo_costos.json dentro del repo."""
    for cand in [
        repo / "costos_unitarios" / "catalogo_costos.json",
        repo / "catalogo_costos.json",
    ]:
        if cand.exists():
            return cand
    # búsqueda recursiva
    for cand in repo.glob("**/catalogo_costos.json"):
        return cand
    return None


# ---------------------------------------------------------------------------
# Cliente
# ---------------------------------------------------------------------------

class ClienteCosteo:
    """Consume el sistema de costos PONTEM (buscador oficial o JSON directo)."""

    def __init__(self, repo: Optional[str | Path] = None):
        self.repo: Optional[Path] = None
        self.indice_path: Optional[Path] = None
        self.buscador = None            # módulo buscador_costos (Nivel 3)
        self._conceptos: Optional[list] = None
        self._por_clave: dict[str, dict] = {}
        self.nivel = None               # 2/3 = buscador oficial, 1 = JSON directo
        self._descubrir(repo)

    # -- descubrimiento -------------------------------------------------
    def _descubrir(self, repo: Optional[str | Path]):
        self.repo = Path(repo).resolve() if repo else descubrir_repo_costeo()
        if self.repo is None:
            return

        self.indice_path = ruta_indice(self.repo)

        # Nivel 3: intentar importar el buscador oficial
        buscador_py = self.repo / "buscador_costos.py"
        if buscador_py.exists():
            try:
                spec = importlib.util.spec_from_file_location(
                    "buscador_costos_pontem", str(buscador_py)
                )
                mod = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)  # type: ignore
                self.buscador = mod
                self.nivel = 3
                return
            except Exception:
                self.buscador = None

        # Nivel 1: leer el JSON directamente
        if self.indice_path and self.indice_path.exists():
            self.nivel = 1
            self._cargar_json()

    # -- carga JSON (Nivel 1) ------------------------------------------
    def _cargar_json(self):
        if self._conceptos is not None:
            return
        data = json.loads(self.indice_path.read_text(encoding="utf-8"))
        self._conceptos = data.get("conceptos", [])
        for c in self._conceptos:
            k = c.get("clave")
            if k:
                self._por_clave[k] = c

    # -- estado ---------------------------------------------------------
    @property
    def disponible(self) -> bool:
        return self.buscador is not None or bool(self._conceptos) or (
            self.indice_path is not None and self.indice_path.exists()
        )

    def descripcion_contrato(self) -> str:
        if self.nivel == 3:
            return f"Nivel 3 — buscador oficial (buscador_costos.py) en {self.repo}"
        if self.nivel == 1:
            return f"Nivel 1 — lectura directa del JSON en {self.indice_path}"
        return "No disponible"

    # -- búsqueda -------------------------------------------------------
    def buscar(self, texto: str, limite: int = 20,
               unidad: Optional[str] = None) -> list[dict]:
        """Busca conceptos por texto. Usa el contrato oficial si está disponible."""
        if self.buscador is not None:
            try:
                return self.buscador.buscar(
                    texto, unidad=unidad, limite=limite,
                    ruta_indice=str(self.indice_path) if self.indice_path else None,
                )
            except TypeError:
                # versión de buscador sin ruta_indice/unidad
                return self.buscador.buscar(texto, limite=limite)
            except Exception:
                pass  # cae al fallback

        # Fallback Nivel 1: búsqueda por fracción de tokens
        self._cargar_json()
        if not self._conceptos:
            return []
        q = _norm(texto)
        tokens = q.split()
        q_uni = _norm(unidad) if unidad else None
        salida = []
        for c in self._conceptos:
            if q_uni and q_uni != _norm(c.get("unidad")):
                continue
            desc = _norm(c.get("descripcion"))
            if not tokens or not all(True for _ in [0]):
                score = 1.0
            hay = sum(1 for t in tokens if t in desc)
            if tokens and hay == 0:
                continue
            score = hay / len(tokens) if tokens else 1.0
            if q and q in desc:
                score += 0.5
            if tokens and hay == len(tokens):
                score += 0.3
            salida.append((score, c))
        salida.sort(key=lambda x: -x[0])
        res = []
        for score, c in salida[:limite]:
            d = dict(c)
            d["score"] = round(score, 3)
            res.append(d)
        return res

    def por_clave(self, clave: str) -> Optional[dict]:
        """Devuelve el concepto con esa clave exacta (o None)."""
        if self.buscador is not None:
            try:
                return self.buscador.buscar_por_clave(
                    clave, ruta_indice=str(self.indice_path) if self.indice_path else None
                )
            except TypeError:
                return self.buscador.buscar_por_clave(clave)
            except Exception:
                pass
        self._cargar_json()
        return self._por_clave.get(clave)

    def todos(self) -> list[dict]:
        """Todos los conceptos (para indexar)."""
        self._cargar_json()
        return self._conceptos or []


# ---------------------------------------------------------------------------
# Test rápido
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    cli = ClienteCosteo()
    print("Repo:", cli.repo)
    print("Contrato:", cli.descripcion_contrato())
    if cli.disponible:
        for r in cli.buscar("muro de block 15cm", limite=3):
            print(f"  [{r.get('score')}] {r.get('clave')} | {r.get('unidad')} | ${r.get('precio_unitario')} | {r.get('descripcion','')[:50]}")