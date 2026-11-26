# ESPECIFICACIÓN FUNCIONAL
## Sistema de presupuestación PONTEM — `presupuestar.py` / `app.py`

**Versión:** 2.0 | 2026-11-26 — alineada con el contrato de consumo v1.2 del
repo de costeos (`C:\Pontem_dev\costeo\dist\contrato_costos\1.2.0\`)

> Changelog de esta versión: el presupuesto pasó de cotizar el PU del catálogo
> por `VLOOKUP` a cotizar el **PU oficial del modelo canónico**
> (`v_concepto_pu.precio_unitario`), y el desglose del APU pasa a venir de la
> BD canónica de costos en modo read-only. Ver §3.

---

### 0. Objetivo

Un sistema que construye un presupuesto de construcción en un único `.xlsx` a
partir de:

- un archivo de entrada con los conceptos a cotizar,
- el **contrato de consumo** del repo de costeos (catálogo JSON + BD canónica),
- una plantilla `.xlsx` que ya trae el FASAR incrustado.

El `.xlsx` resultante contiene fórmulas nativas de Excel: si el usuario modifica
un insumo, una cuadrilla o una variable, se recalculan los APUs y el
presupuesto.

---

### 1. Alcance

El sistema:

1. Pide por CLI (o por la app web) la ruta del archivo de entrada; descubre
   solo el repo de costeos y la plantilla.
2. Copia la plantilla a la carpeta de salida. Nunca modifica el original.
3. **Consume el catálogo** por el contrato publicado (`catalogo_costos.json` del
   paquete `dist/contrato_costos/<versión>/`).
4. **Consume la BD canónica** (`costos_canonico.db`) en modo read-only para el
   desglose de los APUs y para el PU que manda.
5. Lee el archivo de entrada (4 o 5 columnas).
6. Empareja cada concepto de entrada con un concepto del catálogo.
7. Escribe en la copia de la plantilla las hojas `Presupuesto`, `Variables`,
   `Insumos`, `Cuadrillas`, `APU_<id>` y `Revisión` (la Portada y las hojas
   FASAR vienen de la plantilla).
8. Valida el `.xlsx` construido (9 verificaciones).
9. Guarda un único archivo de salida y su PDF.

---

### 2. Entradas del sistema

**Archivo de entrada (`.xlsx`)**

Cuatro columnas exactas: `id | concepto | unidad | cantidad`, o cinco:
`id | partida | concepto | unidad | cantidad`. En el formato de 4 columnas se
aceptan filas de encabezado `PARTIDA` / `SUBPARTIDA` incrustadas (columna `id`
que empieza con `PARTIDA`/`SUB`, o `id` con texto y las demás columnas vacías):
agrupan los conceptos siguientes.

**Repo de costeos**

Se **descubre solo** (carpetas hermanas, priorizando el que tiene
`buscador_costos.py` y el catálogo) o se fuerza con `PONTEM_COSTOS_DB` /
`PONTEM_COSTOS_JSON`. No se assume ninguna ruta fija.

**Plantilla (`plantillas/machote.xlsx`)**

Define estilos, layout, Portada y **trae el FASAR ya incrustado** (`Parametros`,
`Calculo_categoria`, `Variables`). Ya no existe un archivo FASAR externo: la
hoja `Parametros` queda oculta en el libro final.

---

### 3. Contrato de consumo del repo de costeos

El consumidor se ata al **artefacto publicado**, no al interior del repo:

| Superficie | Para qué | Dónde |
|---|---|---|
| `catalogo_costos.json` | clave → precio unitario | `dist/contrato_costos/<versión>/` |
| `buscador_costos.py` | búsqueda de candidatos (nivel 3) | raíz del repo de costeos |
| `costos_canonico.db` (read-only) | APU, rendimientos, origen del precio | `dist/contrato_costos/<versión>/` |
| `v_concepto_pu` | **el PU que manda** y su origen | vista pública |
| `v_apu_resumen`, `v_apu_linea_costo` | desglose valuado del APU | vistas públicas |

Reglas que el sistema respeta:

- La BD canónica se abre **read-only** (`?mode=ro`) y solo se consultan **vistas
  públicas**; nunca tablas internas.
- **No** se leen los Excel de `PU2026/` ni las tablas planas de `costos.db`.
- Se respeta la **prioridad del PU del modelo canónico**:
  `publicado > APU > manual`. El PU cotizado es `v_concepto_pu.precio_unitario`.
- Si la canónica no trae el APU de un concepto, el **desglose** se completa con
  el índice local de matrices (fuente *no contractual*, marcada `matrices` en la
  Revisión); el PU sigue siendo el oficial. Se desactiva con
  `PONTEM_COSTOS_SIN_MATRICES=1`.

---

### 4. Salida del sistema

Un `.xlsx` en la carpeta de salida (`salidas/<timestamp>/` en la app web), con
las hojas:

1. `Portada` (de la plantilla)
2. `Presupuesto`
3. `Parametros` (de la plantilla, **oculta**)
4. `Calculo_categoria` (de la plantilla)
5. `Cuadrillas`
6. `Variables`
7. `Insumos`
8. `APU_<id>` (una por concepto con desglose)
9. `Revisión`

Regla de valores duros: solo los **PU oficiales** del catálogo (columna F del
Presupuesto) y los **costos unitarios de la BD canónica** que el usuario puede
editar. Todo lo demás es fórmula.

---

### 5. Flujo de ejecución

```
$ python presupuestar.py "entrada/conceptos.xlsx"

Contrato: Nivel 3 — buscador oficial en C:\Pontem_dev\costeo
BD canónica de costos: 838 APUs (12,927 conceptos)

MATCHEO DE CONCEPTOS
[1] Matcheando conceptos...        ✓ 48 conceptos
[2] Construyendo presupuesto...
[1/8] Copiando plantilla...
[2/8] Escribiendo Variables...
[3/8] Preparando APUs...
[4/8] Escribiendo Cuadrillas (IDs locales)...
[5/8] Escribiendo Insumos (materiales)...
[6/8] Creando hojas APU...
[7/8] Escribiendo Presupuesto...
[8/8] Escribiendo Portada...
Guardando...                        ✓ ./presupuesto.xlsx
Validando libro...                  ✓ 0 errores, 0 advertencias
[3] Generando PDF (Presupuesto + APUs)...
```

La app web (Streamlit) ofrece además la casilla **"Cotizar el PU del APU"**;
la CLI, la opción `--cotizar-apu`.

---

### 6. Hoja `Presupuesto`

| Col | Campo | Tipo |
|---|---|---|
| A | Id | texto (id de entrada, ej. `CIM-01`) |
| B | Clave PONTEM | texto (clave global del catálogo) |
| C | Descripción | texto |
| D | Cantidad | número |
| E | Unidad | texto |
| F | P.U. | **valor** (PU oficial) o fórmula `='<APU>'!F<n>` (modo APU) |
| G | Importe | fórmula `=D*F` |
| H | % | (participación; se reserva) |

Encabezado de partida (`PARTIDA: <nombre>`) y subtotales por partida. Al pie:
costo directo, indirectos (campo + oficina), financiamiento, utilidad, IVA y
gran total, todo como fórmulas encadenadas.

**El PU de la columna F es el oficial del catálogo** (`v_concepto_pu.precio_unitario`).
Solo si el usuario elige el modo APU se escribe la referencia a la hoja APU, de
modo que editar `Insumos` mueva el total.

---

### 7. Hoja `Variables`

Parámetros globales editables, con columna de notas.

| Variable | Origen |
|---|---|
| UMA diaria, tope SBC, ISN | FASAR (plantilla) |
| Indirectos campo / oficina | editables |
| Financiamiento, utilidad, IVA | editables |
| Días de trabajo por mes | 30 |
| Rendimiento default | 1 |

### 8. Hoja `Calculo_categoria` (de la plantilla)

Réplica de la calculadora FASAR: `Categoría | Sn | SDI | SBC con tope | E&M
cuota fija | E&M excedente | Cargas % SBC | Total cargas | Ps | FASAR | Sr -
Salario real`. Sus fórmulas leen `Variables`.

---

### 9. Hoja `Cuadrillas`

| ID | Cuadrilla | Cant. 1 | Categoría 1 | Cant. 2 | Categoría 2 | Cant. 3 | Categoría 3 | $/Jornal | Descripción |
|---|---|---|---|---|---|---|---|---|---|

`$/Jornal` (col I) es fórmula que suma, por renglón, cantidad × `Sr` de la
categoría en `Calculo_categoria`.

**IDs locales:** la hoja usa IDs propios (`CUAD001`…), no las claves del repo.
La columna `$/Jornal` es editable: es el punto de extensión para el costo de
mano de obra mientras el modelo canónico no traiga la composición de las
cuadrillas (ver §16).

---

### 10. Hoja `Insumos`

Tabla maestra única: materiales, equipo, herramienta, cuadrillas, categorías y
conceptos del repo.

| Col | Campo | Tipo |
|---|---|---|
| A | Id local | texto único (`INS-001`…) |
| B | Tipo | `material`, `equipo`, `herramienta`, `cuadrilla`, `categoria`, `concepto_repo` |
| C | Descripción | texto |
| D | Unidad | texto |
| E | Costo unitario | valor editable o fórmula (cuadrilla → `Cuadrillas!I`) |
| F | Cantidad total | número (Σ cantidad_APU × cantidad_obra) |
| G | Importe | fórmula `=E*F` |
| H | Notas | texto |
| I | Clave origen | **clave global** del repo (trazabilidad) |

**IDs locales + clave global:** el `VLOOKUP` de las hojas APU liga por el ID
local (columna A), que es único; la columna I conserva la clave global del
catálogo para poder rastrearla.

---

### 11. Hoja `APU_<id>`

Una hoja por concepto con desglose. El **id de entrada** va en `C1` (es lo que
permite al validador localizar la hoja aunque el PU sea un valor).

| Col | Campo | Tipo |
|---|---|---|
| A | Id local | texto (`INS-005`) |
| B | Insumo | texto |
| C | Unidad | texto |
| D | Cantidad | número |
| E | Costo unit. | fórmula `=IFERROR(VLOOKUP(A6,Insumos!$A$5:$G$11,5,FALSE),<costo canónico>)` |
| F | Importe | fórmula `=D*E` |
| G | Clave origen | clave global |

Secciones con subtotal: `Materiales`, `Mano de Obra`, `Equipo`, `Herramienta`,
`Auxiliares`.

Al pie: costo directo, indirectos, financiamiento, utilidad, **PU DE VENTA**
(`=F<cd>+F<ind>+F<util>`) y **PU OFICIAL (CATALOGO)** (valor, con la nota del
origen del precio). Si el desglose no vino de la canónica, la nota lo marca
como fuente no contractual.

---

### 12. Hoja `Revisión`

Siempre presente. 15 columnas (18 si hubo sugerencias de IA):

| # | Columna | Contenido |
|---|---|---|
| 1–10 | Id entrada, Concepto entrada, Unidad, Cantidad, Clave sugerida, Descripción sugerida, Score, Método, Estado, Desglosar APU | matcheo |
| 11 | PU oficial | `v_concepto_pu.precio_unitario` |
| 12 | PU APU | `v_concepto_pu.precio_unitario_apu` |
| 13 | Dif. % | `(PU APU − PU oficial) / PU oficial`, resaltado si > 1 % |
| 14 | Origen precio | `publicado` / `apu` / `manual` |
| 15 | Fuente desglose | `canonica` / `matrices` (resaltada si no es canónica) |
| 16–18 | Clave IA, Confianza IA, Motivo IA | solo si se usó IA |

Estados: `confirmado`, `dudoso`, `sin_match`. `Desglosar APU` la llena el
usuario con S o N.

---

### 13. Hoja `Portada` (de la plantilla)

Propuesta económica con la tabla de conceptos (id, concepto, unidad, cantidad,
P.U.) y subtotales por partida; el validador comprueba que esa tabla sume el
costo directo del `Presupuesto`.

---

### 14. Dependencias entre hojas

```
Variables ──► Calculo_categoria ──► Cuadrillas ──┐
                                                 ├──► Insumos ──► APU_<id> ──► Presupuesto
Insumos (materiales, equipo, herramienta) ───────┘
BD canónica ──► APU_<id> (cantidades y costos) ──► Presupuesto (PU oficial)
```

Reglas:

- `Presupuesto` no lee `Insumos` por `VLOOKUP`: su descripción y unidad vienen
  del catálogo y su PU es el oficial (o una referencia a la hoja APU).
- `APU_<id>` solo lee `Insumos`.
- `Cuadrillas` solo lee `Calculo_categoria`.
- `Calculo_categoria` solo lee `Variables`.
- Editar un material en `Insumos` recalcula las hojas APU que lo usan; en modo
  APU eso mueve el PU del `Presupuesto`.

---

### 15. Módulos

| Módulo | Entrada | Salida |
|---|---|---|
| `costeo_client.py` | repo de costeos | catálogo publicado + `buscar()` (nivel 3) |
| `apu_db.py` | BD canónica (read-only) | APU por clave + PU oficial y su origen |
| `apu_source.py` | matrices `.xlsx` | **fallback** de desglose (no contractual) |
| `lector_entrada.py` | `.xlsx` de entrada | conceptos + partidas/subpartidas |
| `matcher.py` | conceptos de entrada e índice | mapeo con score y método |
| `llm.py` | candidatos dudosos | top-3 sugerencias con justificación |
| `constructor.py` | mapeo confirmado + plantilla | copia con hojas y fórmulas |
| `validador.py` | `.xlsx` construido | reporte de errores y advertencias |
| `pdf_export.py` | `.xlsx` | PDF (Presupuesto + APUs) |
| `formulas.py`, `estilos.py`, `visual_dna.py` | — | helpers de fórmulas y estilo |

`indexador.py` quedó como índice local para los scripts de prueba; el flujo de
producción usa `costeo_client.py`.

---

### 16. Política del LLM

- Se invoca solo si el usuario lo activa y el score fuzzy está bajo el umbral.
- Recibe la descripción de entrada y los 10 mejores candidatos.
- Devuelve un top-3 con justificación corta; nunca inventa claves.
- Nunca escribe en `Presupuesto`, `Insumos` ni las hojas APU: solo en `Revisión`.
- Siempre hay confirmación humana antes de construir.

---

### 17. Validaciones antes de guardar

1. Toda clave usada en un `VLOOKUP` a `Insumos`/`Cuadrillas` existe.
2. Toda clave de las hojas APU existe en `Insumos`.
3. Los **IDs locales** de `Insumos` son únicos.
4. Toda categoría usada en `Cuadrillas` existe en `Calculo_categoria`.
5. Consistencia de unidades entre `Presupuesto`, las hojas APU e `Insumos`.
6. Ausencia de ciclos de referencia entre hojas.
7. Sintaxis válida de las fórmulas.
8. Cadena de totales y cuadre aritmético (recálculo con Excel si está
   disponible); ninguna cuadrilla en `$0`; la Portada suma el costo directo.
9. El **PU cotizado coincide con el PU oficial**: si difiere, se reporta como
   **advertencia con el impacto en dinero** (se omite si la BD canónica no está
   disponible).

Si hay errores, el archivo se guarda además como `presupuesto_con_errores.xlsx`.

---

### 18. Dependencias Python

```
openpyxl>=3.1
pandas>=2.0
rapidfuzz>=3.0
pydantic>=2.0
typer>=0.12
rich>=13.0
pyyaml>=6.0
formulas>=1.2
anthropic>=0.30
```

---

### 19. Estructura de carpetas

```
Presupuestos-excel/
├─ presupuestar.py            # CLI
├─ app.py                     # app web (Streamlit)
├─ config.yaml
├─ requirements.txt
├─ src/
│  ├─ apu_db.py               # BD canónica (read-only)
│  ├─ apu_source.py           # fallback de matrices (local, no contractual)
│  ├─ costeo_client.py        # catálogo publicado
│  ├─ lector_entrada.py
│  ├─ matcher.py
│  ├─ llm.py
│  ├─ constructor.py
│  ├─ validador.py
│  ├─ pdf_export.py
│  ├─ formulas.py / estilos.py / visual_dna.py
│  └─ indexador.py            # índice local (scripts de prueba)
├─ plantillas/machote.xlsx    # plantilla + FASAR incrustado
├─ cache/                     # índice local de matrices (NO versionado)
├─ salidas/
└─ tests/
```

---

### 20. Deuda técnica y pendientes

Ver [`BACKLOG.md`](BACKLOG.md) para el seguimiento vigente. Los puntos abiertos
que afectan a esta especificación:

- El modelo canónico todavía **no trae la composición de las cuadrillas**
  (`cuadrilla_integrantes` y `categorias_mo` están vacías): por eso el
  `$/Jornal` de la hoja `Cuadrillas` es editable a mano. Cuando el repo de
  costeos las pueble, el costo de la mano de obra pasa a ser calculado.
- La canónica solo trae **838 de 9,348 APUs** y captura el 63.9 % de las
  líneas: el resto se completa con el fallback de matrices (marcado como fuente
  no contractual). Requerimiento detail en
  `C:\Pontem_dev\costeo\docs\REQUERIMIENTOS_CARRIL_APUS.md`.
- `docs/plan.md` es **bitácora histórica**: describe el estado del proyecto en
  fechas pasadas y no se actualiza.
