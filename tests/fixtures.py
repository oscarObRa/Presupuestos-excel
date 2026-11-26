"""
fixtures.py
===========
Datos de prueba AUTOCONTENIDOS para la suite de pruebas.

No dependen del repo de costos (C:\\Pontem_dev\\costeo) ni de las matrices:
todo se genera en memoria, de modo que las pruebas se puedan ejecutar en CI
o en cualquier equipo.

Uso:
    from tests.fixtures import ClienteFake, APUFake, ENTRADA, PROYECTO, construir_de_prueba
"""

import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

# ---------------------------------------------------------------------------
# Catálogo de prueba
# ---------------------------------------------------------------------------
# CAT-003 y CAT-004 tienen la MISMA descripción en unidades distintas:
# es el caso que producía precios 1000× mayores (KG vs TON).
CATALOGO = [
    {"clave": "CAT-001", "descripcion": "Trazo y nivelacion de terreno",
     "unidad": "M2", "precio_unitario": 500.0},
    {"clave": "CAT-002", "descripcion": "Muro de block hueco de concreto 15x20x40cm",
     "unidad": "M2", "precio_unitario": 350.0},
    {"clave": "CAT-003", "descripcion": "Acero de refuerzo fy=4200 kg/cm2",
     "unidad": "KG", "precio_unitario": 45.0},
    {"clave": "CAT-004", "descripcion": "Acero de refuerzo fy=4200 kg/cm2",
     "unidad": "TON", "precio_unitario": 45000.0},
    {"clave": "CAT-005", "descripcion": "Excavacion de material B",
     "unidad": "M3", "precio_unitario": 149.69},
    {"clave": "CAT-006", "descripcion": "Cimiento de concreto f'c=150",
     "unidad": "M3", "precio_unitario": 1800.0},
    {"clave": "CAT-007", "descripcion": "Pintura latex sobre muros", "unidad": "M2",
     "precio_unitario": 95.0},
]

# APU de prueba para CAT-002 (materiales + una cuadrilla)
# Los campos de trazabilidad del precio reproducen el modelo canonico de
# costos: `pu_oficial` es el PU que manda (publicado > APU > manual) y
# `pu_apu` es el que calcula el APU.
APU_CAT002 = {
    "clave": "CAT-002",
    "descripcion": "Muro de block hueco de concreto 15x20x40cm",
    "unidad": "M2",
    "pu": 350.0,
    "pu_oficial": 350.0,
    "pu_apu": 372.5,
    "origen_precio": "publicado",
    "diferencia_publicado_apu": 22.5,
    "costo_directo": 2004.56,
    "fuente": "canonica",
    "insumos": [
        {"tipo": "material", "clave": "MAT-001",
         "descripcion": "Block de concreto 15x20x40", "unidad": "PZA",
         "cantidad": 52.5, "costo": 20.48, "total": 1075.2},
        {"tipo": "material", "clave": "MAT-002",
         "descripcion": "Cemento gris", "unidad": "KG",
         "cantidad": 100.0, "costo": 3.5, "total": 350.0},
        # Cuadrilla: su clave general NO existe en la hoja Cuadrillas de la
        # plantilla (solo hay CUAD001...), por eso el libro debe darle un ID
        # local y escribirlo en esa hoja.
        {"tipo": "mano_obra", "clave": "1A1P",
         "descripcion": "CUADRILLA No 5 (1 ALBAÑIL + 1 PEON)", "unidad": "JOR",
         "cantidad": 0.25, "costo": 2317.42, "total": 579.36},
    ],
}

ENTRADA = [
    {"id": "CIM-01", "partida": "CIMENTACION",
     "concepto": "Trazo y nivelacion de terreno", "unidad": "M2", "cantidad": 100},
    {"id": "CIM-02", "partida": "CIMENTACION",
     "concepto": "Muro de block hueco de concreto 15x20x40cm", "unidad": "M2",
     "cantidad": 200},
    {"id": "CIM-03", "partida": "CIMENTACION",
     "concepto": "Excavacion de material B", "unidad": "M3", "cantidad": 50},
]

PROYECTO = {
    "nombre": "PRUEBA AUTOMATICA",
    "clave": "TEST-01",
    "cliente": "Cliente de prueba",
    "fecha": "2026-11-25",
    "cotizacion": "COT-001",
    "revision": "R01",
    "indirectos_campo": 0.14,
    "indirectos_oficina": 0.05,
    "financiamiento": 0.0135,
    "utilidad": 0.08,
    "iva": 0.16,
}


class ClienteFake:
    """ClienteCosteo mínimo: todos() / buscar() / por_clave()."""

    def __init__(self, conceptos=None):
        self._conceptos = [dict(c) for c in (conceptos or CATALOGO)]
        self._por_clave = {c["clave"]: c for c in self._conceptos}

    def todos(self):
        return [dict(c) for c in self._conceptos]

    def buscar(self, texto, limite=20, unidad=None):
        q = str(texto or "").strip().lower()
        res = []
        for c in self._conceptos:
            if q and q in str(c.get("descripcion", "")).lower():
                if unidad and str(c.get("unidad", "")).upper() != str(unidad).upper():
                    continue
                res.append(dict(c))
        return res[:limite]

    def por_clave(self, clave):
        c = self._por_clave.get(clave)
        return dict(c) if c else None


class APUFake:
    """Proveedor de APUs mínimo para el constructor.

    Reproduce el contrato del proveedor canonico: `apu_de()` devuelve el
    desglose y la trazabilidad del precio, y `concepto()` el PU oficial.
    """

    def __init__(self, apus=None):
        self._apus = dict(apus or {"CAT-002": APU_CAT002})

    @property
    def disponible(self):
        return True

    def apu_de(self, clave, paquete=None):
        a = self._apus.get(clave)
        if a is None:
            return None
        a = dict(a)
        a.setdefault("pu_oficial", a.get("pu"))
        a.setdefault("pu_apu", a.get("pu"))
        a.setdefault("origen_precio", "manual")
        a.setdefault("diferencia_publicado_apu", 0.0)
        a.setdefault("fuente", "canonica")
        return a

    def concepto(self, clave):
        a = self._apus.get(clave)
        if not a:
            return None
        return {
            "clave": clave,
            "precio_unitario": a.get("pu_oficial", a.get("pu")),
            "precio_unitario_apu": a.get("pu_apu", a.get("pu")),
            "origen_precio": a.get("origen_precio", "manual"),
            "diferencia_publicado_apu": a.get("diferencia_publicado_apu", 0.0),
            "apu_id": 1 if a.get("insumos") else None,
        }

    def cargar_indice(self):
        return len(self._apus)


def matcher_de_prueba(cliente=None, **kwargs):
    """Matcher con el catálogo de prueba."""
    from src.matcher import Matcher
    return Matcher(cliente=cliente or ClienteFake(), **kwargs)


def construir_de_prueba(carpeta_salida, validar=False, recalcular=False,
                        entrada=None, apus=None, incluir_revision=False,
                        usar_pu_oficial=True):
    """Construye un presupuesto de prueba y devuelve (ruta, resultados, matcher)."""
    from src.constructor import construir_presupuesto

    cliente = ClienteFake()
    matcher = matcher_de_prueba(cliente)
    resultados = matcher.emparejar(entrada if entrada is not None else ENTRADA)
    ruta = construir_presupuesto(
        ruta_plantilla=str(RAIZ / "plantillas" / "machote.xlsx"),
        proyecto=PROYECTO,
        resultados_matcheo=resultados,
        carpeta_salida=str(carpeta_salida),
        usar_timestamp=False,
        apu_provider=APUFake(apus),
        incluir_revision=incluir_revision,
        validar=validar,
        recalcular=recalcular,
        usar_pu_oficial=usar_pu_oficial,
    )
    return ruta, resultados, matcher
