"""
indexador.py
============
Índice de búsqueda sobre costos.db (SQLite) - reemplaza el Indexador basado en JSON.

Uso:
    from src.indexador import Indexador

    idx = Indexador(r"C:\Pontem_dev\costeo\costos.db")
    idx.cargar()

    # Búsqueda por clave exacta
    concepto = idx.por_clave("10401-061")

    # Búsqueda por texto en descripción
    resultados = idx.buscar("muro de block", limite=10)

    # Todos los conceptos
    todos = idx.todos()
"""

import json
import os
import re
import sqlite3
import unicodedata
from pathlib import Path
from typing import Optional, List, Dict


class Indexador:
    """Índice de búsqueda sobre costos.db (SQLite)."""

    def __init__(self, ruta_db: str):
        self.ruta_db = Path(ruta_db)
        self._db_conn: Optional[sqlite3.Connection] = None
        self._cargado = False
        self._total = 0

    def _get_conn(self) -> sqlite3.Connection:
        if self._db_conn is None:
            self._db_conn = sqlite3.connect(str(self.ruta_db))
            self._db_conn.row_factory = sqlite3.Row
        return self._db_conn

    def cargar(self) -> int:
        """Carga el índice desde la BD SQLite. Retorna el número de conceptos."""
        if not self.ruta_db.exists():
            raise FileNotFoundError(f"No se encontró la BD: {self.ruta_db}")

        # Verificar que las tablas existen
        cur = self._get_conn().cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='conceptos'")
        if not cur.fetchone():
            raise RuntimeError("Tabla 'conceptos' no existe en la BD")

        cur.execute("SELECT COUNT(*) FROM conceptos")
        self._total = cur.fetchone()[0]
        self._cargado = True
        return self._total

    @property
    def cargado(self) -> bool:
        return self._cargado

    @property
    def total(self) -> int:
        return self._total

    def _get_conn(self) -> sqlite3.Connection:
        if self._db_conn is None:
            self._db_conn = sqlite3.connect(str(self.ruta_db))
            self._db_conn.row_factory = sqlite3.Row
        return self._db_conn

    def por_clave(self, clave: str) -> Optional[dict]:
        """Busca un concepto por su clave exacta. Retorna None si no existe."""
        if not self._cargado:
            raise RuntimeError("El índice no está cargado. Llama a .cargar() primero.")
        cur = self._get_conn().execute(
            "SELECT * FROM conceptos WHERE clave = ?", (clave,)
        )
        row = cur.fetchone()
        return dict(row) if row else None

    def buscar(self, texto: str, limite: int = 10) -> list[dict]:
        """Búsqueda simple por subcadena en descripción (case-insensitive)."""
        if not self._cargado:
            raise RuntimeError("El índice no está cargado. Llama a .cargar() primero.")

        texto_lower = texto.lower()
        cur = self._get_conn().execute(
            "SELECT * FROM conceptos WHERE lower(descripcion) LIKE ? LIMIT ?",
            (f"%{texto.lower()}%", limite)
        )
        return [dict(row) for row in cur.fetchall()]

    def todos(self) -> list[dict]:
        """Retorna todos los conceptos."""
        if not self._cargado:
            raise RuntimeError("El índice no está cargado. Llama a .cargar() primero.")
        cur = self._get_conn().execute("SELECT * FROM conceptos")
        return [dict(row) for row in cur.fetchall()]

    def stats(self) -> dict:
        """Estadísticas del índice."""
        if not self._cargado:
            return {"cargado": False, "total": 0}

        cur = self._get_conn().cursor()
        cur.execute("SELECT COUNT(DISTINCT paquete) FROM conceptos")
        paquetes = cur.fetchone()[0]
        cur.execute("SELECT COUNT(DISTINCT unidad) FROM conceptos")
        unidades = cur.fetchone()[0]
        cur.execute("SELECT COUNT(DISTINCT aplicacion) FROM conceptos")
        aplicaciones = cur.fetchone()[0]

        return {
            "cargado": True,
            "ruta": str(self.ruta_db),
            "total": self._total,
            "paquetes": paquetes,
            "unidades": unidades,
            "aplicaciones": aplicaciones,
        }


if __name__ == "__main__":
    idx = Indexador(r"C:\Pontem_dev\costeo\costos.db")
    idx.cargar()
    print(f"Total: {idx.total}")
    print(idx.buscar("muro", limite=3))