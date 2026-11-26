# BACKLOG — Sistema de presupuestación PONTEM

**Creado:** 2026-11-26 · **Última actualización:** 2026-11-26 (B-12/B-13/B-14, R-6..R-9)
**Alcance:** `C:\Pontem_dev\Presupuestos-excel` (este repo) y los requerimientos
que este repo le hace a `C:\Pontem_dev\costeo`.

Este archivo es el **registro de seguimiento**. Si algo no está en el README, en
la especificación funcional o aquí, no existe.

---

## Estados

| Estado | Significado |
|---|---|
| `PENDIENTE` | no empezado |
| `EN CURSO` | empezado, sin cerrar |
| `HECHO` | cerrado y verificado |
| `BLOQUEADO` | depende de algo externo (otro repo, una decisión) |
| `REQUIERE COSTEO` | el trabajo es del repo `C:\Pontem_dev\costeo` |

---

## 1. Resumen

| # | Ítem | Estado | Notas |
|---|---|---|---|
| B-01 | Consumir la BD canónica y cotizar el PU oficial | `HECHO` | commit `4b34e22` (antes `d75b205`) |
| B-02 | Alinear README y mensajes con el contrato | `HECHO` | commit `e884698` |
| B-03 | Sacar `cache/apu_index.json` de git + invalidación | `HECHO` | commit `6cdbce9` |
| B-04 | Limpiar el historial del blob de 12.3 MB | `HECHO` | `git filter-repo` + force push |
| B-05 | Actualizar la documentación (spec + consumo) | `HECHO` | `docs/especificacion-funcional.md` v2.0, `docs/consumo-costos.md` |
| B-06 | Limpiar scripts obsoletos de la raíz | `HECHO` | ver §3 |
| B-07 | Scripts sueltos en `tests/` con rutas absolutas | `HECHO` | movidos a `scripts/diagnostico/` |
| B-08 | `src/indexador.py`: solo lo usan scripts de prueba | `PENDIENTE` | decidir si se mueve |
| B-09 | Descartes silenciosos en `src/apu_source.py` | `PENDIENTE` | medido: 0.03 %, no material |
| B-10 | Purgar objetos no referenciados en GitHub | `PENDIENTE` | ver §5 (riesgos) |
| B-11 | Commitear el requerimiento en el repo de costeos | `PENDIENTE` | ver §4 |
| B-12 | **Bidireccionalidad: devolver cambios a costeo** | `PENDIENTE` | diseño cerrado; requiere R-6..R-8 (ver §3) |
| B-13 | Exportador de propuestas (`src/retorno_costeo.py`) | `PENDIENTE` | depende de R-6/R-7 |
| B-14 | Consumir la composición de cuadrillas de la canónica | `PENDIENTE` | depende de R-9 (ver §3) |

---

## 2. Hecho (contexto para no perderlo)

### B-01 · Consumo canónico y PU oficial — `HECHO`

El consumidor se ató al artefacto publicado (`dist/contrato_costos/<versión>/`):

- `src/apu_db.py` reescrito: BD canónica **read-only** (`?mode=ro`), solo vistas
  públicas (`v_concepto_pu`, `v_apu_resumen`, `v_apu_linea_costo`). **Cero
  lecturas de `costos.db`**.
- El PU cotizado es `v_concepto_pu.precio_unitario` (prioridad del modelo:
  publicado > APU > manual). Antes el libro cotizaba el PU recalculado por la
  hoja APU, que difería del oficial en **803 de 838** conceptos con APU.
- Modo alterno `--cotizar-apu` (CLI) / casilla en la app: cotiza el PU
  recalculado, para que editar `Insumos` mueva el total.
- Trazabilidad visible: Revisión con `PU oficial`, `PU APU`, `Dif. %`,
  `Origen precio`, `Fuente desglose`; hoja APU con `PU OFICIAL (CATALOGO)`.
- Validador: check **9/9** (el PU cotizado debe coincidir con el oficial; si
  difiere, advertencia con impacto en dinero).

**Impacto medido:** con el PU oficial, `10401-172` cotiza $52.97 en vez de
$791.58; `10502-081`, $100.33 en vez de $49,812.

### B-02 · README y mensajes — `HECHO`

`README.md`, `src/indexador.py`, `src/constructor.py` y el mensaje de arranque
de la CLI ya no mencionan `costos.db` ni `APU_<clave>`; el flujo documentado
son los 8 pasos reales.

### B-03 · Cache local con invalidación — `HECHO`

`cache/apu_index.json` (12.3 MB, derivado de `G:\Mi unidad\Costos\PU2026\`)
salió del control de versiones. `src/apu_source.py` lo invalida con una huella
= hash sha1 de (ruta, mtime, tamaño) de cada Excel, guardada en
`cache/apu_index.meta.json`.

Se descartó el criterio "mtime máximo": la prueba mostró que editar una matriz
que **no** sea la más reciente no lo detectaba y dejaba el cache vencido en
silencio.

Medido: 231 s la primera vez (contenido idéntico al que estaba versionado),
~1 s después, regenera al tocar cualquier matriz.

### B-04 · Limpieza del historial — `HECHO`

`git filter-repo --path cache/apu_index.json --invert-paths --force` +
force push con `--force-with-lease`. `.git` pasó de 3.8 MB a 2.8 MB; el blob ya
no existe en ninguna rama. Respaldo previo en
`C:\Users\Oscar Obregon\AppData\Local\Temp\opencode\presupuestos-excel-respaldo-pre-filter.bundle`.

**Los SHAs cambiaron** (reescritura de historia): `d75b205` → `4b34e22`,
`41af682` → `e884698`, `fc77808` → `6cdbce9`.

### B-05 · Documentación — `HECHO`

- `docs/especificacion-funcional.md` → v2.0, reescrita contra el sistema real
  (hojas, columnas, fórmulas, PU oficial vs APU, IDs locales, 9 validaciones).
- `docs/consumo-costos.md` → reescrito contra el contrato v1.2 (5 niveles; se
  quitan los niveles que el contrato prohíbe). Todos los ejemplos de API se
  **ejecutaron y verificaron**.
- `README.md` → contrato canónico, qué PU se cotiza, el cache local.

---

## 3. Pendiente en este repo

### B-06 · Limpieza de scripts obsoletos — `HECHO`

| Archivo | Qué era | Qué se hizo |
|---|---|---|
| `fix_insumos.py` | parche one-off que reescribía `src/constructor.py` con `str.replace`; ya aplicado | `git rm` |
| `fix_insumos_v2.py` | ídem, segunda versión | `git rm` |
| `test_run.py` | prueba manual de la CLI con stdin simulado; superseded por `tests/run_all.py` | `git rm` |
| `test_match.py` | script de depuración del matcheo (untracked) | borrado |
| `test_match2.py` | ídem (untracked) | borrado |
| `test_match3.py` | ídem (untracked) | borrado |

Verificado antes de borrar: `tests/run_all.py` solo corre 3 suites y ningún
módulo de `src/` ni el workflow de CI importaban estos archivos.

### B-07 · Scripts sueltos en `tests/` — `HECHO`

`tests/` tenía **68 archivos versionados**; la suite real son 5
(`run_all.py`, `fixtures.py` y las 3 suites). El resto se movió:

| Destino | Qué |
|---|---|
| `scripts/diagnostico/codigo/` | 29 scripts de inspección y depuración |
| `scripts/diagnostico/datos/` | 10 catálogos y entradas de ejemplo |
| `scripts/diagnostico/salidas/` | 22 salidas generadas (PDF, PNG, `.xlsx`, logs) — **ignorada por git** |

Los scripts que importan `src.*` recibieron un bootstrap
(`sys.path.insert(0, str(Path(__file__).resolve().parents[3]))`) y sus rutas
relativas se actualizaron de `tests/...` a `scripts/diagnostico/...`.

**Advertencia documentada en `scripts/diagnostico/README.md`:** estos scripts
**pueden estar desactualizados** (usan rutas absolutas a `G:\`, `H:\` y a la
API anterior). Por ejemplo `test_paso3.py` pide una clave `formato` que
`validar_entrada()` no devuelve: bug **preexistente**, no causado por el
movimiento. No son un contrato.

Se añadieron `tests/README.md` y `scripts/diagnostico/README.md`.

### B-08 · `src/indexador.py` — `PENDIENTE`

Solo lo usan scripts de diagnóstico (`scripts/diagnostico/codigo/diag_top.py`,
`debug_scoring.py`, etc.); el flujo de producción usa `src/costeo_client.py`.
Decidir: moverlo a `scripts/diagnostico/` o dejarlo con la nota que ya tiene
en el docstring ("queda como índice local para los scripts de prueba y
diagnóstico").

### B-09 · Descartes silenciosos en `src/apu_source.py` — `PENDIENTE` (bajo)

`src/apu_source.py:199-201` descarta una línea cuando el encabezado de sección
no se reconoce, **sin avisar**. Medido sobre las 428 matrices reales: **20
líneas de 65,745 (0.03 %)**. No es material, pero instrumentar un contador
evita que el número crezca sin que nadie lo note.

### B-12 · Bidireccionalidad: devolver cambios al repo de costeos — `PENDIENTE`

**Pedido del usuario (2026-11-26):** el consumo no puede ser de una sola vía. Si
durante la presupuestación aparece algo nuevo (un precio, un concepto, una
cuadrilla, un rendimiento), debe poder copiarse de regreso a
`C:\Pontem_dev\costeo`.

#### Decisión de diseño: el consumidor NO escribe en la BD canónica

Dos razones, ambas verificables:

- La BD canónica es **regenerable**: `actualizar_catalogo.py --con-migracion`
  la reconstruye desde cero y **borra los cambios manuales**
  (`C:\Pontem_dev\costeo\costos_unitarios\herramientas\actualizar_catalogo.py:84-91`).
- El contrato vigente lo prohíbe textualmente: *"❌ Escribir en la BD canónica"*
  (`C:\Pontem_dev\costeo\docs\CONTRATO_CONSULTA_COSTOS_SQL.md:250`).

#### Patrón elegido: artefacto de propuesta (JSON) + importador en costeo

Es el mismo patrón que ya funciona en sentido contrario:
`C:\Pontem_dev\costeo\costos_unitarios\herramientas\importar_catalogo_externo.py`
importa un JSON de catálogo externo, hace UPSERT por clave en `conceptos` y
deja rastro en `migracion_log`. El consumidor produciría ese JSON; **costeo lo
aplicaría**. Así el retorno es explícito, auditable, reversible y no depende de
que el consumidor tenga acceso de escritura a la BD.

#### Qué se devuelve (payload), por prioridad

| Sección | Qué | Dónde vive hoy en el libro | Tabla destino en la canónica |
|---|---|---|---|
| `conceptos[]` | conceptos nuevos o con PU corregido | hoja `Presupuesto` (filas `sin_match`) y hoja `Revisión` | `conceptos.precio_unitario_manual` |
| `insumos[]` | costo unitario corregido de un material o equipo | hoja `Insumos` col E (valor editable) + col I (clave global) | `insumo_precios` (con `fecha_vigencia_desde`) |
| `cuadrillas[]` | `$ / Jornal` y composición por categoría | hoja `Cuadrillas` col I (editable) | `cuadrilla_integrantes` + `categorias_mo.salario_base_diario` |
| `hallazgos[]` | divergencias y faltantes, sin cambiar precios | hoja `Revisión` (Dif. %, Origen precio, Fuente desglose) + vistas de diagnóstico | ninguno: es informational |

**No se devuelve:** cantidades de obra (son del proyecto, no del catálogo), ni
el libro completo, ni APUs corregidos (ver trampa 2).

#### Detección de qué cambió el usuario (detalle de diseño)

El libro **no guarda el costo original**: materiales, equipos y herramientas se
escriben como valor editable en `Insumos` col E (`src/constructor.py:448-451`),
y las cuadrillas como fórmula `VLOOKUP` a `Cuadrillas!I`
(`src/constructor.py:443-447`). Por eso el exportador debe **re-leer el `.xlsx`
ya editado** y comparar contra la canónica (usando la clave global de la col
I) o contra el índice de matrices cuando `fuente="matrices"`. El retorno es un
paso **posterior** a la edición del libro, no parte de la construcción.

#### Trampas ya detectadas

1. **El PU devuelto no mandaría.** Con la prioridad `publicado > APU > manual`
   (`C:\Pontem_dev\costeo\sql\README.md:334-335`), un PU manual devuelto solo
   aplica a conceptos **sin APU**. Para los 838 con APU hay que limpiar
   `precio_unitario_publicado` (congelado) o registrar un override: es una
   decisión de costeo (R-8). Contexto: hoy **806 de 838** conceptos con APU
   tienen un PU publicado que difiere >1 % del calculado por su APU.
2. **No hay importador de APU.** `importar_catalogo_externo.py` solo importa
   conceptos con precio manual, sin desglose. Devolver un APU corregido exige
   un importador nuevo: fuera del alcance de la primera versión.
3. **La composición de cuadrillas ya existe en la BD viva** (363 filas en
   `cuadrilla_integrantes`, 33 en `categorias_mo`) pero **no está publicada** y
   este repo no la consulta: solo usa la tasa consolidada de la línea de APU.
   Publicarla es R-9; consumirla es B-14.

#### Orden de implementación (propuesta)

1. **R-6** (importador en costeo) + **R-7** (contrato v1.3 con el formato del
   payload): sin contrato no hay retorno confiable.
2. **B-13** (exportador en este repo) + test.
3. **R-8** (política de prioridad) antes de que se devuelva el primer PU.
4. **R-9** (publicar 1.3.0 con cuadrillas y consumos) y **B-14**.

### B-13 · Exportador de propuestas (`src/retorno_costeo.py`) — `PENDIENTE`

- `exportar_propuesta(ruta_xlsx, ruta_json)`: lee el libro editado y arma el
  JSON con `meta` (formato, versión, proyecto, fecha, origen) + las 4 secciones.
- CLI: `presupuestar.py --exportar-propuesta <salida.json>`.
- App web: botón "Exportar propuestas a costeo" junto al PDF.
- Validación del payload antes de escribir (mismo criterio que el importador de
  costeo) y resumen de lo que se propone cambiar.
- Test en `tests/` apoyándose en el fixture actual (`tests/fixtures.py`).

### B-14 · Consumir la composición de cuadrillas de la canónica — `PENDIENTE`

Hoy `src/constructor.py` arma la hoja `Cuadrillas` desde la **tasa consolidada**
de la línea de APU (`src/constructor.py:316-370`), no desde
`v_cuadrilla_detalle` / `v_cuadrilla_costo_jornal`. Con la BD viva ya poblada
(363 integrantes, 33 categorías) se puede poblar la hoja con la composición real
(categoría y cantidad por integrante) en vez de una tasa opaca, y con eso el
`$/Jornal` deja de ser una caja negra. Depende de **R-9** (publicar esa
información).

---

## 4. Requerimientos al repo `C:\Pontem_dev\costeo` — `REQUIERIE COSTEO`

Documento completo (con métricas verificadas):
**`C:\Pontem_dev\costeo\docs\REQUERIMIENTOS_CARRIL_APUS.md`** — *sin commitear
en ese repo* (ver B-11).

| # | Requerimiento | Impacto medido |
|---|---|---|
| R-1 | El loader debe tolerar el nombre de la hoja (`MATRICES` **y** `Sheet`) | +7,396 APUs |
| R-2 | Las claves de insumo sin guion deben resolverse, no descartarse | +30,984 líneas (51.8 %) |
| R-3 | El parser no debe depender de offsets fijos de columna | 595 de 838 APUs con CD mal |
| R-4 | Poblar `cuadrilla_integrantes` y `categorias_mo` | **hecho en la BD viva** (363 y 33); falta publicarlo (R-9) y consumirlo (B-14) |
| R-5 | `costos.db` → snapshot de la canónica; borrar `esquema.sql` (0 KB) | elimina la ambigüedad |
| R-6 | Importador de propuestas de retorno: `costos_unitarios/herramientas/importar_propuesta.py` (JSON → UPSERT en `conceptos` / `insumo_precios` / `cuadrilla_integrantes`, con validación, `--dry-run` y registro en `migracion_log`) | habilita B-13 |
| R-7 | Publicar la superficie de retorno en el contrato (versión **1.3**): formato del payload + CLI del importador | el consumidor necesita un contrato, no una convención |
| R-8 | Política de prioridad para un PU devuelto (limpiar el publicado congelado / registrar override / solo alertar) | hoy un PU devuelto **no mandaría**: `publicado > APU > manual` |
| R-9 | Publicar una versión nueva (1.3.0) con lo que ya está en la BD viva y no en el snapshot 1.2.0: `cuadrilla_integrantes` 363, `categorias_mo` 33, `equipo_consumos` 105, y 14,541 conceptos vs 12,927 | el consumidor usa el snapshot por defecto → hoy no ve cuadrillas ni consumos |

**Contexto:** el catálogo está completo (12,927 conceptos en el snapshot
publicado; 14,541 en la BD viva, 0 colisiones); el hueco es el carril de APUs
(838 de 9,348) y las cuadrillas.

**Prioridad para el repo de costeos (propuesta):**
R-3 > R-2 > R-1 > R-8 > R-6 > R-7 > R-9 > R-4 > R-5.

> Nota: R-6, R-7, R-8 y R-9 son **nuevos** (2026-11-26) y nacen del requisito de
> bidireccionalidad (B-12). No están todavía en
> `C:\Pontem_dev\costeo\docs\REQUERIMIENTOS_CARRIL_APUS.md`, que sigue
> enfocado en el carril de APUs.

---

## 5. Riesgos y decisiones vigentes

### Decisiones

| Decisión | Motivo |
|---|---|
| `usar_pu_oficial=True` por defecto | es la regla del modelo canónico; el libro cotizaba distinto del oficial en 803 de 838 |
| `--cotizar-apu` / casilla en la app | quien quiera recálculo vivo lo pide explícitamente |
| Fallback de matrices **solo para desglose**, marcado `matrices` | cubre 8,510 conceptos sin APU en la canónica, sin contaminar el PU |
| IDs locales + clave global de trazabilidad | el `VLOOKUP` liga por ID local (único); la clave global queda para rastrear |
| El `$/Jornal` de `Cuadrillas` es editable a mano | la canónica publicada no trae la composición de las cuadrillas (R-4/R-9) |
| **El retorno a costeo es un artefacto JSON, no escritura directa en la canónica** | la BD es regenerable (`--con-migracion` borra cambios manuales) y el contrato prohíbe escribir en ella |
| **El exportador lee el `.xlsx` editado, no el modelo en memoria** | el dato nuevo lo deja el usuario en el libro; el libro no guarda el costo original, hay que compararlo contra la canónica |
| **El consumidor aplica, costeo decide** | cada cambio de precio pasa por revisión humana en el repo dueño del modelo |

### Riesgos

| Riesgo | Impacto | Mitigación |
|---|---|---|
| GitHub conserva los objetos no referenciados tras el force push (hasta ~90 días) | el blob de 12.3 MB puede seguir accesible por SHA | B-10: pedir purga a soporte de GitHub |
| La canónica solo trae 838 de 9,348 APUs | sin acceso a `G:\Mi unidad\Costos\PU2026\` se pierde el desglose fallback | hasta que R-1..R-3 cierren |
| El `$/Jornal` se edita a mano | error humano en el costo de mano de obra | el validador avisa si una cuadrilla queda en $0 |
| **806 de 838 APUs** (BD viva) tienen PU publicado que difiere >1 % del APU | si el repo de costeos limpia los publicados, **los PU cambian** | avisar antes de ejecutar `migrar_costos_db.py --sin-concelar` |
| `cache/` se regenera en ~231 s la primera vez | la primera ejecución con fallback es lenta | es local y automático |
| El snapshot publicado (1.2.0) va por detrás de la BD viva | el consumidor no ve 363 integrantes de cuadrilla, 33 categorías, 105 consumos de equipo ni 1,614 conceptos extra | R-9: publicar 1.3.0 |
| Un PU devuelto por el consumidor no tendría efecto | `publicado > APU > manual` deja al precio congelado por encima | R-8: definir la política antes del primer retorno |
| El retorno sin contrato se vuelve un acuerdo informal | cada consumidor inventaría su propio formato | R-7: publicarlo en el contrato v1.3 |

---

## 6. Cómo se mantiene

- Al cerrar un ítem: moverlo a §2 (Hecho) con el commit y la métrica.
- Al abrir uno nuevo: añadirlo a §1 y a §3 (o §4 si es del repo de costeos).
- Los números de este backlog se **miden**, no se estiman: si cambia el
  contrato publicado (v1.3.0), hay que re-medir y actualizar §4.
- Las cifras de la canónica se miden sobre **las dos** bases, porque difieren:
  el snapshot publicado (`C:\Pontem_dev\costeo\dist\contrato_costos\1.2.0\costos_canonico.db`,
  12,927 conceptos) y la BD viva
  (`C:\Pontem_dev\costeo\costos_canonico.db`, 14,541 conceptos). El consumidor
  usa el snapshot por defecto (verificado: `src/apu_db.py::_descubrir_db`).
