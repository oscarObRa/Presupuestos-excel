"""
apu_db.py
=========
Proveedor de APUs que consume `costos.db` (SQLite) según el contrato del repo
de costos:

    conceptos   -> catálogo (clave, descripción, unidad, PU)
    apu_lineas  -> desglose (categoria, insumo_id, descripcion, cantidad,
                   unidad, costo_unitario, importe)
    apu_totales -> totales (tipo, valor, pct)

`apu_de(clave)` devuelve la misma estructura que MatricesProvider:
    {"clave","descripcion","unidad","pu","insumos":[{tipo,clave,descripcion,unidad,cantidad,costo,total}]}

Si la BD no tiene el APU de un concepto, cae a las MATRICES en Excel (fallback),
para no perder cobertura.
"""

from __future__ import annotations

import os
import sqlite3
from pathlib import Path
from typing import Optional


def _descubrir_db(repo: Optional[str] = None) -> Optional[Path]:
    env = os.environ.get("PONTEM_COSTOS_DB")
    if env and Path(env).exists():
        return Path(env)
    if repo:
        p = Path(repo) / "costos.db"
        if p.exists():
            return p
    # buscar entre carpetas hermanas del presupuestador
    try:
        from src.costeo_client import descubrir_repo_costeo
        r = descubrir_repo_costeo()
        if r:
            p = Path(r) / "costos.db"
            if p.exists():
                return p
    except Exception:
        pass
    for cand in [Path(r"C:/Pontem_dev/costeo/costos.db")]:
        if cand.exists():
            return cand
    return None


# Mapear categoría de la BD a los tipos que usa el constructor
_MAP_CAT = {
    "material": "material",
    "mano_obra": "mano_obra",
    "equipo": "equipo",
    "maquinaria": "maquinaria",
    "auxiliar": "auxiliar",
    "herramienta": "herramienta",
}


class DBProvider:
    """Lee APUs desde costos.db."""

    def __init__(self, repo: Optional[str] = None, db_path: Optional[str] = None):
        self.db_path = Path(db_path) if db_path else _descubrir_db(repo)
        self._conn: Optional[sqlite3.Connection] = None

    def _c(self) -> sqlite3.Connection:
        if self._conn is None:
            self._conn = sqlite3.connect(str(self.db_path))
            self._conn.row_factory = sqlite3.Row
        return self._conn

    @property
    def disponible(self) -> bool:
        return self.db_path is not None and self.db_path.exists()

    def contar_lineas(self) -> int:
        if not self.disponible:
            return 0
        try:
            return self._c().execute("SELECT COUNT(*) FROM apu_lineas").fetchone()[0]
        except Exception:
            return 0

    def apu_de(self, clave: str, paquete: Optional[str] = None) -> Optional[dict]:
        if not self.disponible:
            return None
        try:
            cur = self._c()
            con = cur.execute("SELECT * FROM conceptos WHERE clave = ?", (str(clave),)).fetchone()
            if con is None:
                return None
            idc = con["id"]
            lineas = cur.execute(
                "SELECT categoria, insumo_id, descripcion, cantidad, unidad, "
                "costo_unitario, importe FROM apu_lineas WHERE id_concepto = ? "
                "ORDER BY orden, categoria, id",
                (idc,),
            ).fetchall()
            if not lineas:
                return None
            totales = cur.execute(
                "SELECT tipo, valor, pct FROM apu_totales WHERE id_concepto = ? ORDER BY orden",
                (idc,),
            ).fetchall()

            insumos = []
            for l in lineas:
                cat = _MAP_CAT.get((l["categoria"] or "").strip().lower(), "material")
                cant = float(l["cantidad"] or 0)
                costo = float(l["costo_unitario"] or 0)
                imp = l["importe"]
                imp = float(imp) if imp is not None else round(cant * costo, 4)
                insumos.append({
                    "tipo": cat,
                    "clave": l["insumo_id"],
                    "descripcion": l["descripcion"] or "",
                    "unidad": l["unidad"] or "",
                    "cantidad": cant,
                    "costo": costo,
                    "total": imp,
                })

            # PU: buscar total de venta; si no, usar el PU del concepto
            pu = float(con["precio_unitario"] or 0)
            for t in totales:
                tipo_n = (t["tipo"] or "").lower()
                if "venta" in tipo_n or tipo_n in ("pu", "precio unitario", "precio_unitario"):
                    pu = float(t["valor"] or 0)
                    break

            return {
                "clave": con["clave"],
                "descripcion": con["descripcion"],
                "unidad": con["unidad"],
                "pu": pu,
                "insumos": insumos,
            }
        except Exception:
            return None

    def stats(self) -> dict:
        return {"db": str(self.db_path) if self.db_path else None,
                "lineas": self.contar_lineas(),
                "disponible": self.disponible}


class ProveedorAPU:
    """BD primero; si un concepto no está poblado, usa las MATRICES (Excel)."""

    def __init__(self, repo: Optional[str] = None, db_path: Optional[str] = None):
        self.db = DBProvider(repo=repo, db_path=db_path)
        self._mat = None
        self._repo = repo

    @property
    def disponible(self) -> bool:
        return self.db.disponible or self._matrices().disponible

    def _matrices(self):
        if self._mat is None:
            from src.apu_source import MatricesProvider
            self._mat = MatricesProvider(repo=self._repo)
        return self._mat

    def cargar_indice(self) -> int:
        """Devuelve cuántos conceptos tienen APU (BD + matrices cacheadas)."""
        n = 0
        if self.db.disponible:
            try:
                n = self.db._c().execute(
                    "SELECT COUNT(DISTINCT id_concepto) FROM apu_lineas"
                ).fetchone()[0]
            except Exception:
                n = 0
        # no forzamos matrices aquí (se usan como fallback bajo demanda)
        return n

    def apu_de(self, clave: str, paquete: Optional[str] = None) -> Optional[dict]:
        a = self.db.apu_de(clave, paquete) if self.db.disponible else None
        if a and a.get("insumos"):
            return a
        m = self._matrices()
        if m.disponible:
            return m.apu_de(clave, paquete)
        return a

    def stats(self) -> dict:
        return {"db": self.db.stats(), "matrices": self._matrices().stats() if self._mat else None}


if __name__ == "__main__":
    p = DBProvider()
    print("DB:", p.db_path, "| lineas:", p.contar_lineas())
    if p.disponible:
        row = p._c().execute("SELECT clave FROM conceptos LIMIT 1").fetchone()
        if row:
            a = p.apu_de(row["clave"])
            print("Ejemplo", row["clave"], "insumos:", len(a["insumos"]) if a else 0)