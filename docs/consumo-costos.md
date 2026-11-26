# Cómo consumir el sistema de costos PONTEM desde otros repos

**Actualizado:** 2026-11-26 — alineado con el contrato publicado **v1.2**
(`C:\Pontem_dev\costeo\dist\contrato_costos\1.2.0\`).

> **Regla del contrato** (`LEEME.md` del paquete publicado):
> *"El consumidor se ata al artefacto, no al interior del repo Costos. No leas
> Excel, carpetas fuente ni tablas internas."*
>
> Es decir: consume `dist/contrato_costos/<versión>/` (o las variables de
> entorno), nunca `costos.db`, ni `costos_unitarios/`, ni los Excel de
> `PU2026\`.

---

## Nivel 1 — Leer el JSON publicado (sin dependencias)

```python
import json

with open(r"C:\Pontem_dev\costeo\dist\contrato_costos\1.2.0\catalogo_costos.json") as f:
    data = json.load(f)

# 12,927 conceptos indexados
for c in data["conceptos"]:
    if "muro" in c["descripcion"].lower():
        print(c["clave"], c["precio_unitario"])
```

**Ventaja:** cero dependencias; funciona en cualquier lenguaje.

---

## Nivel 2 — Instalar el paquete pip (API de Python)

```bash
pip install pontem-costos
```

```python
from pontem_costos import buscar, buscar_por_clave, meta

resultados = buscar("muro de block", limite=5)
# -> [{"clave": "14203-015", "precio_unitario": 1234.56, "score": 0.71, ...}, ...]

concepto = buscar_por_clave("10502-081")
print(concepto["precio_unitario"], concepto["unidad"])
```

**Ventaja:** API limpia, 1 línea.

---

## Nivel 3 — Copiar el buscador (stdlib-only)

`buscador_costos.py` solo usa bibliotecas estándar (`json`, `re`, `pathlib`).
Está en el paquete publicado; lo copias a tu repo y funciona:

```python
from buscador_costos import buscar, buscar_por_clave, cargar, meta

# carga manual (el propio `buscar` ya cachea)
_meta, conceptos = cargar(r"c:\...\dist\contrato_costos\1.2.0\catalogo_costos.json")

resultados = buscar("concreto", limite=10,
                     ruta_indice=r"c:\...\dist\contrato_costos\1.2.0\catalogo_costos.json")
```

**Ventaja:** portátil, sin `pip install`, sin acceso a GitHub.

---

## Nivel 4 — La BD canónica (SQLite read-only) ⭐

Para APU, rendimientos, cuadrillas, costo horario y trazabilidad del precio.
Está en el paquete publicado como `costos_canonico.db`.

```python
import sqlite3

uri = r"file:///C:/Pontem_dev/costeo/dist/contrato_costos/1.2.0/costos_canonico.db?mode=ro"
con = sqlite3.connect(uri, uri=True)
con.row_factory = sqlite3.Row

# El PU que manda y su origen (publicado > APU > manual)
for r in con.execute("""
    SELECT clave, descripcion, unidad, precio_unitario, origen_precio
    FROM v_concepto_pu WHERE clave = ?
""", ("10502-081",)):
    print(dict(r))

# La matriz APU valuada
for r in con.execute("""
    SELECT orden, insumo_clave, descripcion, tipo, cantidad_calculada,
           costo_unitario, importe, origen_costo
    FROM v_apu_linea_costo
    WHERE apu_id = (SELECT apu_id FROM v_concepto_pu WHERE clave = '10502-081')
    ORDER BY orden
"""):
    print(dict(r))
```

O con la API que trae el paquete (son **funciones de módulo**, no una clase).
Ojo con la forma del retorno: `obtener_concepto()` y `obtener_apu()` devuelven
un dict con la clave anidada:

```python
from consultas_sql import (
    meta, buscar, obtener_concepto, obtener_apu,
    obtener_rendimientos, obtener_equipo, conectar,
)

DB = r"C:\Pontem_dev\costeo\dist\contrato_costos\1.2.0\costos_canonico.db"

print(meta(DB))
# {'esquema': {...}, 'conceptos': 12927, 'apus': 838, 'lineas_apu': 3982, ...}

concepto = obtener_concepto("10502-081", DB)["concepto"]
print(concepto["clave"], concepto["precio_unitario"], concepto["origen_precio"])

apu = obtener_apu("10502-081", DB)
print(apu["apu"]["costo_directo"], len(apu["lineas"]))
for l in apu["lineas"]:
    print(l["orden"], l["insumo_clave"], l["descripcion"], l["cantidad_calculada"])
```

O con SQL directo (la forma más flexible):

```python
import sqlite3

uri = r"file:///C:/Pontem_dev/costeo/dist/contrato_costos/1.2.0/costos_canonico.db?mode=ro"
con = sqlite3.connect(uri, uri=True)
con.row_factory = sqlite3.Row

for r in con.execute("SELECT clave, precio_unitario, origen_precio FROM v_concepto_pu"):
    print(dict(r))
```

**Reglas:** abrir **read-only** (`?mode=ro`) y consultar **solo vistas
públicas** (`v_concepto_publicado`, `v_concepto_pu`, `v_apu_resumen`,
`v_apu_linea_costo`, `v_rendimientos`, `v_equipo_costo_hora_detalle`,
`v_cuadrilla_detalle`…); nunca tablas internas.

**Ventaja:** es el único nivel que da el desglose del APU y la razón del
precio.

---

## Nivel 5 — Integración desacoplada (copia periódica)

El patrón que usa Retaining Wall: un proceso del repo de costeos escribe un
JSON propio y el otro repo solo lee ese JSON, sin saber que costeo existe.

```python
# En el repo de costeos (no en el consumidor)
data = json.load(open("catalogo_costos.json", encoding="utf-8"))
json.dump({"costos": [c for c in data["conceptos"] if c["aplicacion"] == "general"]},
          open(r"C:\...\Retaining_wall_app\src\data\costos.json", "w", encoding="utf-8"),
          ensure_ascii=False)
```

**Ventaja:** sincronización automática (Task Scheduler) y desacople total.

---

## Retorno de cambios (bidireccionalidad) — EN DISEÑO, no disponible

> ⚠️ **Hoy no existe.** El contrato vigente (v1.2) dice explícitamente
> *"❌ Escribir en la BD canónica"*
> (`C:\Pontem_dev\costeo\docs\CONTRATO_CONSULTA_COSTOS_SQL.md:250`). No abras
> la BD en modo escritura ni intentes un UPSERT desde un consumidor: la BD es
> **regenerable** y `--con-migracion` borra los cambios manuales.

El diseño que se está adoptando (ver [`BACKLOG.md`](BACKLOG.md) B-12 y los
requerimientos R-6..R-9) es el inverso del nivel 5:

```text
Consumidor  --(JSON de propuesta)-->  Repo de costeos  --(importador)-->  BD canónica
```

- El consumidor **produce** un JSON de propuesta (precios corregidos,
  conceptos nuevos, cuadrillas, hallazgos) y **no** escribe en la BD.
- El repo de costeos lo **aplica** con un importador propio, con validación,
  `--dry-run` y registro en `migracion_log`.

Ya existe el precedente en sentido contrario:
`C:\Pontem_dev\costeo\costos_unitarios\herramientas\importar_catalogo_externo.py`
importa un JSON de catálogo externo con UPSERT por clave en `conceptos`.

El formato del payload se publicará como **contrato v1.3**. Hasta entonces no
hay superficie estable: no construyas contra esto.

---

## Variables de entorno (contrato v1.2)

| Variable | Para qué |
|---|---|
| `PONTEM_COSTOS_DB` | ruta exacta de la BD canónica (snapshot o BD viva) |
| `PONTEM_COSTOS_JSON` | ruta exacta del catálogo JSON |

---

## Resumen: ¿Cuál usar según tu caso?

| Si necesitas... | Usa |
|---|---|
| Buscar precios desde cualquier lenguaje | Nivel 1 (JSON directo) |
| API limpia en Python con scoring | Nivel 2 (`pip install`) |
| Portabilidad máxima, sin instalar nada | Nivel 3 (copiar buscador) |
| APU, rendimientos, cuadrillas o costo horario | **Nivel 4 (BD canónica)** |
| Sincronización automática periódica | Nivel 5 (copia + scheduler) |

---

## Nota: qué queda dentro del repo de costeos

Estos artefactos son **internos** del repo de costeos y **no** son superficie
de consumo (el contrato los prohíbe explícitamente); algunos se retirarán:

| Ruta interna | Nota |
|---|---|
| `costos.db` | base legacy de tablas planas, con totales congelados |
| `esquema.sql` | vacío (0 KB) |
| `_apu_loader.py`, `db_costos.py` | pipeline legacy de carga |
| `indice_pu.json` | índice de navegación de archivos Excel de proyecto |
| `CATALOGO_CONSOLIDADO_*.xlsx` | consolidado de trabajo |
| `costos_unitarios/PU2026\*.xlsx` | matrices fuente |

El consumidor **no** debe leerlos. El equivalente contractual de cada uno está
en el paquete publicado (`dist/contrato_costos/<versión>/`).
