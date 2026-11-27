"""
apu_db.py
=========
Proveedor de APUs sobre la **BD canónica de costos** (`costos_canonico.db`),
según el contrato de consumo v1.2 del repo de costeos
(`docs/CONTRATO_CONSULTA_COSTOS_SQL.md` y el `LEEME.md` del paquete publicado).

Reglas del contrato que este módulo respeta:

- La BD se abre **read-only** (`?mode=ro`).
- Solo se consultan **vistas públicas**; nunca tablas internas:

      v_concepto_pu        -> PU oficial y su origen (publicado / APU / manual)
      v_concepto_publicado -> contrato estable de consumo
      v_apu_resumen        -> CD, indirectos, financiamiento, utilidad
      v_apu_linea_costo    -> matriz APU valuada (cantidad o rendimiento)
      v_rendimientos        -> rendimiento por línea

- **No** se leen los Excels de `PU2026/` ni las tablas planas de `costos.db`:
  el contrato lo prohíbe textualmente ("No leas Excel, carpetas fuente ni
  tablas internas").

El PU que manda es `v_concepto_pu.precio_unitario`, que aplica la prioridad
del modelo canónico:

    precio_unitario_publicado  >  APU con líneas  >  precio_unitario_manual

Mientras exista un PU publicado, el APU es **referencia**; por eso
`apu_de()` expone aparte `pu_oficial`, `pu_apu` y `diferencia_publicado_apu`
para que el libro pueda revelar la diferencia en vez de ocultarla.

Estructura que devuelve `apu_de(clave)`:

    {
      "clave", "descripcion", "unidad",
      "pu":            <PU oficial>,
      "insumos": [{"tipo","clave","descripcion","unidad","cantidad",
                    "costo","total","rendimiento","modo_cantidad",
                    "origen_costo"}],
      "pu_oficial", "pu_apu", "origen_precio", "diferencia_publicado_apu",
      "costo_directo", "fuente"
    }

Si la canónica no trae el APU de un concepto, `ProveedorAPU` puede completar
el **desglose** con el índice de matrices (fuente *no contractual*, se marca
`fuente="matrices"`), pero el PU **siempre** sale de la canónica. Se desactiva
con `PONTEM_COSTOS_SIN_MATRICES=1`.
"""

from __future__ import annotations

import os
import sqlite3
from pathlib import Path
from typing import Optional

# Tipos de línea de la canónica -> secciones que usa el constructor.
# `servicio` es la "herramienta menor" migrada como importe fijo.
_MAP_TIPO = {
    "material": "material",
    "mano_obra": "mano_obra",
    "equipo": "equipo",
    "auxiliar": "auxiliar",
    "servicio": "herramienta",
    "herramienta": "herramienta",
    "maquinaria": "maquinaria",
}


def _version_snapshot(repo: Path) -> Optional[Path]:
    """Localiza el snapshot publicado en `dist/contrato_costos/<version>/`."""
    dist = repo / "dist" / "contrato_costos"
    if not dist.is_dir():
        return None
    versiones = []
    for d in dist.iterdir():
        db = d / "costos_canonico.db"
        if d.is_dir() and db.exists():
            versiones.append((_ordinar_version(d.name), db))
    if not versiones:
        return None
    versiones.sort(key=lambda t: t[0], reverse=True)
    return versiones[0][1]


def _ordinar_version(v: str):
    partes = []
    for p in str(v).replace("-", ".").split("."):
        try:
            partes.append(int(p))
        except ValueError:
            partes.append(0)
    return tuple(partes)


def _descubrir_db(repo: Optional[str | Path] = None) -> Optional[Path]:
    """Encuentra la BD canónica.

    Orden: variable `PONTEM_COSTOS_DB` -> snapshot publicado -> canónica en la
    raíz del repo -> repo de costeo entre carpetas hermanas.
    """
    env = os.environ.get("PONTEM_COSTOS_DB")
    if env and Path(env).exists():
        return Path(env)

    candidatos: list[Path] = []
    if repo:
        r = Path(repo)
        snap = _version_snapshot(r)
        if snap:
            candidatos.append(snap)
        candidatos.append(r / "costos_canonico.db")
    try:
        from src.costeo_client import descubrir_repo_costeo
        r2 = descubrir_repo_costeo()
        if r2:
            snap = _version_snapshot(Path(r2))
            if snap:
                candidatos.append(snap)
            candidatos.append(Path(r2) / "costos_canonico.db")
    except Exception:
        pass
    candidatos.append(Path(r"C:/Pontem_dev/costeo/costos_canonico.db"))

    for c in candidatos:
        if c.exists():
            return c
    return None


class CanonicoProvider:
    """Lectura read-only de la BD canónica por vistas públicas."""

    def __init__(self, repo: Optional[str] = None, db_path: Optional[str] = None):
        self.db_path = Path(db_path) if db_path else _descubrir_db(repo)
        self._conn: Optional[sqlite3.Connection] = None

    # -- conexión ------------------------------------------------------
    def _c(self) -> sqlite3.Connection:
        if self._conn is None:
            if not self.db_path:
                raise RuntimeError("No se encontro la BD canonica de costos")
            uri = self.db_path.as_uri() + "?mode=ro"
            self._conn = sqlite3.connect(uri, uri=True)
            self._conn.row_factory = sqlite3.Row
        return self._conn

    @property
    def disponible(self) -> bool:
        return self.db_path is not None and Path(self.db_path).exists()

    def cerrar(self):
        if self._conn is not None:
            try:
                self._conn.close()
            finally:
                self._conn = None

    # -- consultas -----------------------------------------------------
    def concepto(self, clave: str) -> Optional[dict]:
        """Datos del concepto con el PU que manda y su origen."""
        if not self.disponible:
            return None
        try:
            row = self._c().execute(
                """
                SELECT clave, descripcion, unidad, paquete,
                       precio_unitario, precio_unitario_apu,
                       precio_unitario_publicado, precio_unitario_manual,
                       origen_precio, diferencia_publicado_apu,
                       apu_id, costo_directo, subtotal_1, subtotal_2, utilidad
                FROM v_concepto_pu WHERE clave = ?
                """,
                (str(clave),),
            ).fetchone()
        except Exception:
            return None
        return dict(row) if row else None

    def lineas_de(self, apu_id: int) -> list[dict]:
        """Matriz APU valuada (v_apu_linea_costo)."""
        if not self.disponible:
            return []
        try:
            rows = self._c().execute(
                """
                SELECT orden, insumo_clave, descripcion, tipo, unidad,
                       modo_cantidad, cantidad, rendimiento, cantidad_calculada,
                       factor_desperdicio, costo_unitario, importe,
                       origen_costo, precio_unitario_override, motivo_override
                FROM v_apu_linea_costo
                WHERE apu_id = ? AND apu_es_vigente = 1
                ORDER BY orden
                """,
                (apu_id,),
            ).fetchall()
        except Exception:
            return []
        return [dict(r) for r in rows]

    def apu_de(self, clave: str, paquete: Optional[str] = None) -> Optional[dict]:
        """APU de un concepto con el PU oficial de la canónica."""
        con = self.concepto(clave)
        if not con or not con.get("apu_id"):
            return None
        lineas = self.lineas_de(con["apu_id"])
        if not lineas:
            return None

        insumos = []
        for l in lineas:
            tipo = _MAP_TIPO.get((l.get("tipo") or "").strip().lower(), "material")
            cantidad = l.get("cantidad_calculada")
            if cantidad is None:
                cantidad = l.get("cantidad")
                if cantidad is None and l.get("rendimiento"):
                    try:
                        cantidad = 1.0 / float(l["rendimiento"])
                    except Exception:
                        cantidad = 0.0
            cantidad = float(cantidad or 0)
            costo = float(l.get("costo_unitario") or 0)
            total = l.get("importe")
            total = float(total) if total is not None else round(cantidad * costo, 4)
            insumos.append({
                "tipo": tipo,
                "clave": l.get("insumo_clave") or "",
                "descripcion": l.get("descripcion") or "",
                "unidad": l.get("unidad") or "",
                "cantidad": cantidad,
                "costo": costo,
                "total": total,
                "rendimiento": l.get("rendimiento"),
                "modo_cantidad": l.get("modo_cantidad") or "",
                "origen_costo": l.get("origen_costo") or "",
                "override": l.get("precio_unitario_override") is not None,
            })

        pu_oficial = float(con.get("precio_unitario") or 0)
        return {
            "clave": con.get("clave"),
            "descripcion": con.get("descripcion") or "",
            "unidad": con.get("unidad") or "",
            "pu": pu_oficial,
            "insumos": insumos,
            # trazabilidad del precio (regla del modelo canonico)
            "pu_oficial": pu_oficial,
            "pu_apu": float(con.get("precio_unitario_apu") or 0),
            "origen_precio": con.get("origen_precio") or "",
            "diferencia_publicado_apu": float(con.get("diferencia_publicado_apu") or 0),
            "costo_directo": float(con.get("costo_directo") or 0),
            "fuente": "canonica",
        }

    def stats(self) -> dict:
        if not self.disponible:
            return {"db": None, "disponible": False}
        info = {"db": str(self.db_path), "disponible": True}
        try:
            row = self._c().execute(
                "SELECT COUNT(*) FROM v_concepto_publicado").fetchone()
            info["conceptos"] = row[0]
            row = self._c().execute(
                "SELECT COUNT(*) FROM v_apu_resumen").fetchone()
            info["apus"] = row[0]
            row = self._c().execute(
                "SELECT COUNT(*) FROM v_apu_linea_costo").fetchone()
            info["lineas"] = row[0]
        except Exception:
            pass
        return info


# Alias por compatibilidad con código anterior (la BD ya no es `costos.db`).
DBProvider = CanonicoProvider


class ProveedorAPU:
    """APUs de la canónica; si un concepto no tiene APU, completa el DESGLOSE
    con el índice de matrices (fuente no contractual, marcada en `fuente`).

    El PU nunca sale de las matrices: siempre manda la canónica.
    """

    def __init__(self, repo: Optional[str] = None, db_path: Optional[str] = None):
        self.canonico = CanonicoProvider(repo=repo, db_path=db_path)
        self._mat = None
        self._repo = repo

    # -- compatibilidad con el API viejo --------------------------------
    @property
    def db(self) -> CanonicoProvider:
        return self.canonico

    @property
    def db_path(self):
        return self.canonico.db_path

    @property
    def disponible(self) -> bool:
        return self.canonico.disponible or self._matrices().disponible

    def _matrices(self):
        if self._mat is None:
            from src.apu_source import MatricesProvider
            self._mat = MatricesProvider(repo=self._repo)
        return self._mat

    def _usar_matrices(self) -> bool:
        return os.environ.get("PONTEM_COSTOS_SIN_MATRICES", "").strip() not in ("1", "true", "True")

    def cargar_indice(self) -> int:
        """Cuántos conceptos tienen APU en la canónica."""
        if not self.canonico.disponible:
            return 0
        try:
            return self.canonico._c().execute(
                "SELECT COUNT(*) FROM v_apu_resumen").fetchone()[0]
        except Exception:
            return 0

    def concepto(self, clave: str) -> Optional[dict]:
        return self.canonico.concepto(clave)

    def apu_de(self, clave: str, paquete: Optional[str] = None) -> Optional[dict]:
        oficial = self.canonico.concepto(clave) if self.canonico.disponible else None
        a = self.canonico.apu_de(clave, paquete) if self.canonico.disponible else None
        if a and a.get("insumos"):
            return a

        # Fallback: solo el DESGLOSE. El PU sigue siendo el de la canónica.
        if not self._usar_matrices():
            return a
        m = self._matrices()
        if not m.disponible:
            return a
        am = m.apu_de(clave, paquete)
        if not am or not am.get("insumos"):
            return a
        am = dict(am)
        am["fuente"] = "matrices"
        # El PU que la matriz calcula por sí misma se conserva aparte: es la
        # señal de sanidad del desglose (si la matriz da un PU absurdo frente
        # al oficial, el desglose no es confiable y la línea debe revisarse).
        am["pu_matriz"] = float(am.get("pu") or 0)
        am["costo_directo_matriz"] = float(am.get("costo_directo") or 0)
        if oficial:
            pu_oficial = float(oficial.get("precio_unitario") or 0)
            am["pu_oficial"] = pu_oficial
            am["pu"] = pu_oficial
            am["pu_apu"] = float(oficial.get("precio_unitario_apu") or 0)
            am["origen_precio"] = oficial.get("origen_precio") or ""
            am["diferencia_publicado_apu"] = float(oficial.get("diferencia_publicado_apu") or 0)
            am["costo_directo"] = float(oficial.get("costo_directo") or 0)
        return am

    def stats(self) -> dict:
        return {
            "canonica": self.canonico.stats(),
            "matrices": self._matrices().stats() if self._mat else None,
        }


if __name__ == "__main__":
    p = ProveedorAPU()
    print("BD canonica:", p.db_path)
    print("stats:", p.stats())
    if p.canonico.disponible:
        for clave in ("10502-081", "14101-001"):
            a = p.apu_de(clave)
            if a:
                print(f"\n{clave}: PU oficial={a['pu_oficial']} | PU APU={a['pu_apu']} | "
                      f"origen={a['origen_precio']} | fuente={a['fuente']} | "
                      f"lineas={len(a['insumos'])}")
                for i in a["insumos"][:4]:
                    print(f"    {i['tipo']:10s} {i['clave']:16s} cant={i['cantidad']} "
                          f"x {i['costo']} = {i['total']}")
