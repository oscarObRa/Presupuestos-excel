# BACKLOG — Sistema de presupuestación PONTEM

**Creado:** 2026-11-26 · **Última actualización:** 2026-11-26
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
| B-06 | Limpiar scripts obsoletos de la raíz | `PENDIENTE` | ver §3 |
| B-07 | Scripts sueltos en `tests/` con rutas absolutas | `PENDIENTE` | ver §3 |
| B-08 | `src/indexador.py`: solo lo usan scripts de prueba | `PENDIENTE` | decidir si se mueve |
| B-09 | Descartes silenciosos en `src/apu_source.py` | `PENDIENTE` | medido: 0.03 %, no material |
| B-10 | Purgar objetos no referenciados en GitHub | `PENDIENTE` | ver §5 (riesgos) |
| B-11 | Commitear el requerimiento en el repo de costeos | `PENDIENTE` | ver §4 |

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

### B-06 · Limpieza de scripts obsoletos — `PENDIENTE`

| Archivo | Qué es | Decisión |
|---|---|---|
| `fix_insumos.py` | parche one-off que reescribía `src/constructor.py` con `str.replace`; ya aplicado | eliminar |
| `fix_insumos_v2.py` | ídem, segunda versión | eliminar |
| `test_run.py` | prueba manual de la CLI con stdin simulado; superseded por `tests/run_all.py` | eliminar |
| `test_match.py` | script de depuración del matcheo (untracked) | eliminar |
| `test_match2.py` | ídem (untracked) | eliminar |
| `test_match3.py` | ídem (untracked) | eliminar |

**Al borrar:** `git rm` para los trackeados; los untracked se borran del disco.
Verificar antes que ninguno lo importe el flujo de producción ni la CI.

### B-07 · Scripts sueltos en `tests/` — `PENDIENTE`

`tests/` mezcla la suite real (`test_matcher.py`, `test_constructor.py`,
`test_validador.py`, `run_all.py`) con scripts ad-hoc que usan **rutas
absolutas** a `C:\Pontem_dev\costeo` y no corren en CI:

`diag_top.py`, `debug_scoring.py`, `revisar_partida03.py`, `top_importes.py`,
`test_paso3.py`, `test_pdf_apus.py`, `test_revision_ia.py`,
`test_subpartidas.py`, `test_venado_e2e.py`, `test_llm.py`, `check_db.py`.

Propuesta: moverlos a `scripts/diagnostico/` con rutas relativas, o borrarlos.
Esto también evita que `tests/run_all.py` los descubra por accidente.

### B-08 · `src/indexador.py` — `PENDIENTE`

Solo lo usan scripts de prueba (`tests/diag_top.py`, `tests/debug_scoring.py`,
etc.); el flujo de producción usa `src/costeo_client.py`. Decidir: moverlo a
`tests/` o dejarlo con una nota en el docstring (ya dice que "queda como índice
local para los scripts de prueba y diagnóstico").

### B-09 · Descartes silenciosos en `src/apu_source.py` — `PENDIENTE` (bajo)

`src/apu_source.py:199-201` descarta una línea cuando el encabezado de sección
no se reconoce, **sin avisar**. Medido sobre las 428 matrices reales: **20
líneas de 65,745 (0.03 %)**. No es material, pero instrumentar un contador
evita que el número crezca sin que nadie lo note.

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
| R-4 | Poblar `cuadrilla_integrantes` y `categorias_mo` | 0 filas hoy |
| R-5 | `costos.db` → snapshot de la canónica; borrar `esquema.sql` (0 KB) | elimina la ambigüedad |

**Contexto:** el catálogo está completo (12,927 conceptos, 0 colisiones); el hueco
es el carril de APUs (838 de 9,348) y las cuadrillas.

**Prioridad para el repo de costeos:** R-3 > R-2 > R-1 > R-4 > R-5.

---

## 5. Riesgos y decisiones vigentes

### Decisiones

| Decisión | Motivo |
|---|---|
| `usar_pu_oficial=True` por defecto | es la regla del modelo canónico; el libro cotizaba distinto del oficial en 803 de 838 |
| `--cotizar-apu` / casilla en la app | quien quiera recálculo vivo lo pide explícitamente |
| Fallback de matrices **solo para desglose**, marcado `matrices` | cubre 8,510 conceptos sin APU en la canónica, sin contaminar el PU |
| IDs locales + clave global de trazabilidad | el `VLOOKUP` liga por ID local (único); la clave global queda para rastrear |
| El `$/Jornal` de `Cuadrillas` es editable a mano | la canónica no trae la composición de las cuadrillas (R-4) |

### Riesgos

| Riesgo | Impacto | Mitigación |
|---|---|---|
| GitHub conserva los objetos no referenciados tras el force push (hasta ~90 días) | el blob de 12.3 MB puede seguir accesible por SHA | B-10: pedir purga a soporte de GitHub |
| La canónica solo trae 838 de 9,348 APUs | sin acceso a `G:\Mi unidad\Costos\PU2026\` se pierde el desglose fallback | hasta que R-1..R-3 cierren |
| El `$/Jornal` se edita a mano | error humano en el costo de mano de obra | el validador avisa si una cuadrilla queda en $0 |
| 748 de 838 APUs tienen PU publicado que difiere del APU | si el repo de costeos limpia los publicados, **los PU cambian** | avisar antes de ejecutar `migrar_costos_db.py --sin-concelar` |
| `cache/` se regenera en ~231 s la primera vez | la primera ejecución con fallback es lenta | es local y automático |

---

## 6. Cómo se mantiene

- Al cerrar un ítem: moverlo a §2 (Hecho) con el commit y la métrica.
- Al abrir uno nuevo: añadirlo a §1 y a §3 (o §4 si es del repo de costeos).
- Los números de este backlog se **miden**, no se estiman: si cambia el
  contrato publicado (v1.3.0), hay que re-medir y actualizar §4.
