# Cómo consumir el sistema de costos PONTEM desde otros repos

Cinco niveles de integración, ordenados de más simple a más integrada.

---

## Nivel 1 — Leer el JSON directamente (sin dependencias)

```python
import json

with open("costos_unitarios/catalogo_costos.json") as f:
    data = json.load(f)

# Ya tienes 12,329 conceptos indexados
for c in data["conceptos"]:
    if "muro" in c["descripcion"].lower():
        print(c["clave"], c["precio_unitario"])
```

**Ventaja:** Cero dependencias. Sirve en cualquier lenguaje (JS, Go, Rust).

---

## Nivel 2 — Instalar el paquete pip (API de Python)

```bash
pip install git+https://<token>@github.com/pontem-dev/costeo.git
```

```python
from pontem_costos import buscar, buscar_por_clave, meta

resultados = buscar("muro de block", limite=5)
# -> [{"clave": "14203-015", "precio_unitario": 1234.56, ...}, ...]
```

**Ventaja:** API limpia, 1 línea. Así lo consume Modelador-m7.

---

## Nivel 3 — Copiar el buscador (stdlib-only, ~300 líneas)

El archivo `buscador_costos.py` solo usa bibliotecas estándar de Python (`json`, `re`, `pathlib`). Lo copias a cualquier repo y funciona:

```python
# En tu repo:
from buscador_costos import buscar_en_indice, cargar_indice

indice = cargar_indice("ruta/a/catalogo_costos.json")
resultados = buscar_en_indice(indice, "concreto", limite=10)
```

**Ventaja:** Portátil, no necesita `pip install`, no requiere acceso a GitHub.

---

## Nivel 4 — Usar el índice de proyecto (`indice_pu.json`)

Para investigar archivos Excel de un proyecto específico (OMA, AD01, etc.):

```python
import json

with open("indice_pu.json") as f:
    indice = json.load(f)

# Buscar dónde está la tarjeta APU de PCI-05
concepto = indice["indice_cruzado"]["PCI-05"]
print(concepto["tarjeta_apu"]["archivo"])   # -> nombre del Excel
print(concepto["tarjeta_apu"]["hoja"])      # -> hoja PCI-05
print(concepto["catalogos"])                # -> en qué catálogos aparece
```

**Ventaja:** Navegación instantánea entre archivos, sin abrir Excel.

---

## Nivel 5 — Integración directa (como Retaining Wall)

El script `costos_unitarios/integraciones/alimentar_rw.py` hace esto:

```python
# 1. Lee el índice JSON
with open("catalogo_costos.json") as f:
    data = json.load(f)

# 2. Filtra por aplicación ("general")
conceptos_rw = [c for c in data["conceptos"] if c["aplicacion"] == "general"]

# 3. Escribe el JSON que consume Retaining Wall
with open("C:/.../Retaining_wall_app/src/data/costos.json", "w") as f:
    json.dump({"costos": conceptos_rw}, f)
```

Se ejecuta semanalmente con Windows Task Scheduler. El otro repo solo lee su propio JSON, ni siquiera sabe que costeo existe.

**Ventaja:** Sincronización automática periódica, desacople total entre repos.

---

## Resumen: ¿Cuál usar según tu caso?

| Si necesitas... | Usa |
|---|---|
| Buscar precios desde cualquier lenguaje | Nivel 1 (JSON directo) |
| API limpia en Python con scoring | Nivel 2 (`pip install`) |
| Portabilidad máxima, sin instalar nada | Nivel 3 (copiar buscador) |
| Navegar archivos Excel de proyecto | Nivel 4 (`indice_pu.json`) |
| Sincronización automática periódica | Nivel 5 (script + scheduler) |