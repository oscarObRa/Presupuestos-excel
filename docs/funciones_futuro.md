# Funciones Futuro — Ruta de Evolución del Presupuestador PONTEM

**Actualizado:** 2026-11-27 · Baseline: `v1.0.0` (tag en `main`/`master`)

Este documento registra las ampliaciones acordadas tras el cierre de la v1.0.0.
Todas se trabajan **en ramas** (nada directo a `main`); los bugs de depuración
diaria van como `fix/*` cuando son pequeños y contenidos.

> Convención de ramas acordada:
> - `feature/` funcionalidad nueva · `fix/` correcciones · `test/` pruebas · `docs/` documentación
> - Trabajo con commits por paso y tests en verde; se fusiona a `main` solo con aprobación.

---

## 1. Portada con retícula A:Q del ADN (`feature/portada-reticula-aq`)

**Estado:** pendiente · **Prioridad:** alta (es el pendiente grande declarado)

- Reescribir `escribir_portada` (`src/constructor.py`) dejando la retícula
  legacy de 6 columnas (A:F) y adoptando la retícula de 17 columnas (A:Q) del
  ADN visual, con el `CELL_MAP` de coordenadas fijas.
- Referencia de maquetación: `H:\Mi unidad\ICEL\PUs\CASA VENADO\caratula.xlsx`
  (hoja `PROPUESTA`) — sus anchos de columna coinciden exactamente con los
  `COLUMN_WIDTHS` ya documentados en `docs/ADN_ICEL.md`.
- Coordenadas objetivo (ADN): membrete `B1:D3` (logo) y `M1:Q5` (domicilio/RFC
  con borde inferior verde en fila 5), datos del proyecto en `A7`, `L7`, `A10`,
  `D10`, `A13`, `C13:C15`, `C17`, sección económica en `A20`, encabezados de
  tabla en fila 22 (`A22:B22` Part. | `C22:K22` Concepto | `L22` Unidad |
  `M22` Cantidad | `N22` Precio U. | `O22:Q22` Importe).
- Nota técnica: mientras la Portada siga en 6 columnas, NO aplicar
  `COLUMN_WIDTHS` del ADN a esa hoja (fue la causa de que se viera "echa
  bolas"; hoy tiene sus anchos propios como mitigación).

---

## 2. Revisión de CASA VENADO — trabajo humano pendiente (no es código)

**Estado:** en manos del usuario · **Prioridad:** operativa

- Confirmar/rechazar los **121 dudosos** en la hoja Revision (ahí se ve PU
  oficial vs PU calculado y la diferencia %).
- Resolver los **55 sin_match** a mano o con la IA del sidebar.
- Confirmar/corregir los **outliers marcados en rojo** (PU 0; p.ej.
  subestación de $130k que el matcheo arrastró).

---

## 3. IA para dudosos / sin match (`feature/ia-dudosos`)

**Estado:** existe el gancho (`src/llm.py`, checkbox en la web) · **Prioridad:** media

- Pulir el flujo: resolver dudosos y sin_match con LLM, dejar trazabilidad
  (clave IA, confianza, motivo) en la hoja Revision.
- Definir criterio de confirmación asistida (la IA sugiere, el usuario confirma).

---

## 4. Segunda empresa en el ADN visual (`feature/adn-empresa-2`)

**Estado:** infraestructura lista (`src/visual_dna` + factory) · **Prioridad:** media

- Crear `src/visual_dna/<empresa>.py` heredando de `BaseVisualDNA` (paleta,
  tipografía, assets, `COLUMN_WIDTHS`, `CELL_MAP`, membrete).
- Registrarla en `VisualDNAFactory` y documentar en `docs/ADN_<EMPRESA>.md`.
- El proyecto llega con `proyecto["empresa"] = "clave"` o `empresa=` explícito;
  una clave no registrada lanza `ValueError` (identidad contractual, no se
  maquilla con otra marca).

---

## 5. Empaquetado multiusuario del catálogo (`feature/buscador-exe` + `feature/publicar-sqlitecloud`)

**Estado:** diseño acordado · **Prioridad:** media-alta (requiere 2 ramas coordinadas)

### Contexto

Los usuarios finales no tienen Python/IDE. En lugar del pip (Nivel 2 del
contrato), el catálogo de costeo se sirve desde **SQLite Cloud** y los
consumidores usan un **EXE** (patrón ya probado en
`C:\Pontem_dev\Gestion-Vba\control_obras`: PyInstaller + `installer.iss` +
API keys en `.secrets`).

### Uso real del sistema (aclarado con el usuario)

- **NO es colaboración en el mismo proyecto**: cada usuario trabaja por su
  cuenta y su resultado final es un Excel propio.
- El catálogo es **solo lectura** y ya es inmutable/versionado por contrato.
- Por eso el problema multiusuario se reduce a *distribuir la versión
  vigente del catálogo*; no hay escritura concurrente ni sync de datos.

### Ramas del trabajo

**a) `feature/publicar-sqlitecloud` (repo costeo, lo hace el usuario dueño):**
- Script `publicar_nube.py`: al liberar versión, sube el
  `costos_canonico.db` (~11 MB) como base nueva en la instancia.
- **Disciplina clave: cada versión = una base distinta** (`catalogo_1_2_0`,
  `catalogo_1_3_0`, ...). Nunca pisar la base vigente: así se conserva la
  trazabilidad de "¿con qué precios se cotizó?".
- No activar el motor CloudSync/replicación del panel (evita los triggers por
  tabla documentados en `sync/nube.py` de control_obras): el catálogo es solo
  lectura, no se replica entre pares.
- API keys de solo lectura por usuario; registrar la versión + hash del
  manifest en cada artefacto generado.
- Publicar el canal en el **contrato v1.3** (junto con el diseño de
  bidireccionalidad JSON de propuesta).

**b) `feature/buscador-exe` (consumidor, este repo o costeo):**
- EXE "Buscador PONTEM": núcleo `buscador_costos.py`/`consultas_sql.py`
  (stdlib-only) + UI + driver `sqlitecloud` (ojo: `sqlite3.connect` NO entiende
  `sqlitecloud://`; usar el driver, ver lecciones en `sync/nube.py` de
  control_obras).
- Modo recomendado **híbrido/pull-on-start**: al arrancar consulta la versión
  vigente en la nube; si difiere de la local, descarga y cachea el snapshot;
  trabaja 100% local (rápido y offline). La nube queda como canal de
  distribución, no como runtime.
- Conexión directa a la nube como alternativa (tokens por usuario en
  `.secrets`, nunca en el repo).

### Alternativas de distribución del snapshot (por simplicidad)

| Opción | Cómo | Cuándo |
| :--- | :--- | :--- |
| Snapshot dentro del EXE | `resources/` con JSON + `.db` congelados | Máximo aislamiento y trazabilidad |
| Carpeta de red de solo lectura | `PONTEM_COSTOS_JSON` / `PONTEM_COSTOS_DB` apuntando al share | Cero infraestructura (lectura de SQLite en share es segura; escribir NO) |
| SQLite Cloud con tokens | Driver nativo + API key por usuario | Actualización central sin re-distribuir archivos |

### Cambios en Presupuestos-excel (pequeños, retrocompatibles)

- `ruta_indice()` / `_descubrir_db()` ganan el canal nube: si existen
  `SQLITECLOUD_URL` + `SQLITECLOUD_APIKEY` (o `PONTEM_COSTOS_DB=sqlitecloud://...`),
  leen de la nube; si no, snapshot local. **Offline-first**: la red falla, el
  snapshot local no.
- Imprimir en el libro generado la versión del catálogo usada (ya se muestra
  en la web; falta estamparla en el Excel).
- Distribución sin matrices: para uso oficial canónico, `PONTEM_COSTOS_SIN_MATRICES=1`
  (las matrices `PU2026` son "fuente no contractual" y su ruta no existe en
  máquinas ajenas).

### Advertencias registradas (de control_obras, ya resueltas ahí)

- `sqlite3.connect("sqlitecloud://...")` falla sin diagnóstico: usar el driver
  `sqlitecloud`.
- El `database_id` del panel NO es el nombre de la BD; la URL debe llevar el
  nombre en el path.
- `SQLiteCloudResult` no es iterable ni tiene `.rows`: leer con
  `get_value(fila, col)` vía `nube.consultar()`.
- Nunca correr migraciones DDL destructivas contra la nube.

---

## 6. Otros pendientes menores (del uso real)

- `generar_plantilla.py`: usa `DNA()` sin importarlo ni definirlo (legacy
  roto, no afecta al flujo porque `machote.xlsx` ya existe). Decidir:
  arreglar o retirar.
- Cuadrillas: consolidar criterio de tasas cuando una misma cuadrilla aparece
  con tarifas distintas entre matrices (hoy gana la tasa mayoritaria).
- PDF de la Portada: revisar paginación con la retícula nueva cuando se
  reescriba (Sección 1).

---

## 7. Workbench de revisión y ciclo de enriquecimiento del catálogo

**Estado:** diseño acordado en sesión (2026-11-27) · **Prioridad:** alta — es el salto
de producto: de "generador de Excel" a **estación de trabajo de estimación**.

### 7.1 El cambio de centro de gravedad

Hoy: *matchear → generar libro → abrir Excel → corregir en hoja Revision → re-ejecutar*.
Propuesto: **todo se revisa en el UI del software antes de imprimir**, y el Excel/PDF
es el resultado, no el escritorio.

- Vista *workbench*: lista filtrable por estado (dudoso / sin_match / confirmado),
  partida y score. Al hacer clic en un concepto: candidatos lado a lado con score,
  desglose APU con insumos y costos (editable), unidad, cantidad.
- Acciones por concepto: **marcar confiable / corregir / re-matchear / capturar
  concepto nuevo** (la entrada ya viene semiestructurada; el usuario la redacta como
  concepto de catálogo: descripción, unidad, insumos con rendimientos).
- Señalamientos en vivo del validador (outliers, unidades, divergencia PU oficial
  vs calculado) — todo lo que hoy hace `src/validador.py`, presentado en la sesión.
- **Decisión de entregables: AMBOS exportables siempre.** PDF = entregable formal
  (portada ICEL + presupuesto + APUs, `pdf_export.py` ya existe); Excel = "modo
  vivo" para quien quiera ajustar insumos después (la cadena de fórmulas
  `Insumos → APU → Presupuesto` de `c791aa2` se mantiene para ese modo).
- Al cerrar la sesión: cuadre estático (validador sin recálculo) y export.

### 7.2 El ciclo de enriquecimiento (regla de oro formalizada)

El trabajo de revisión del estimador alimenta el catálogo para todos. **Nada
escribe directo al canónico**: las correcciones aterrizan en una **tabla separada
de acumulación** (`propuestas`), y **costeo es el validador/estandarizador que
publica adecuadamente** al canónico (cola de curaduría + importador con
`--dry-run` y `migracion_log`, como ya está diseñado en `consumo-costos.md`).

| Qué cambió en la revisión | Acción que costeo aplica al canónico | Por qué |
| :--- | :--- | :--- |
| Solo precio de un insumo | **Nueva fila de precio con vigencia** (`insumo_precios.fecha_vigencia_*`), no UPDATE destructivo | "¿Cuánto costaba cuando cotizamos VENADO?" siempre tiene respuesta |
| Estructura del APU (cantidades, insumos, cuadrilla) | **Nueva versión del APU** del mismo concepto (`apus.version` / `es_vigente`) | El concepto no cambia de identidad; su desglose sí |
| Especificación / unidad / descripción | **Concepto NUEVO** con clave nueva; el viejo queda como referencia (deprecable/alias) | Un presupuesto emitido nunca cambia de significado retroactivamente |
| Alcance inexistente | Concepto nuevo capturado desde la entrada, con desglose obligatorio o marcado "PU manual, sin APU" | Evitar referencias huecas en el catálogo |

Reglas de calidad del ciclo:

- Solo viaja **lo que el humano confirmó/corrigió/creó** — el matcher no
  enriquece por sí solo; enriquece la corrección humana.
- Cada contribución lleva **procedencia**: quién, cuándo, en qué presupuesto
  (p.ej. VENADO-R02).
- **Dedup en el importador**: antes de dar de alta un concepto nuevo, correr el
  matcher contra el canónico vigente y señalar coincidencias altas ("esto se
  parece 85% a X") — el estimador decide si es nuevo o el mismo mejor redactado.
- Curaduría: costeo valida/estandariza y publica; los conceptos nuevos y cambios
  de estructura pasan por cola de aprobación. (La auto-aceptación de
  actualizaciones de precio se puede evaluar más adelante, con evidencia fuerte.)

### 7.3 Cómo se integra con la nube (complementa la Sección 5)

Dos bases con roles distintos en SQLite Cloud:

- `catalogo_<versión>` — **solo lectura**, tokens de lectura, inmutable por versión.
- `propuestas` — **escritura** con tokens de contribuidor: acumula las
  correcciones de cada estimador con su procedencia. El importador de costeo la
  consume (dry-run → curaduría → publicación al canónico en la siguiente versión).
- Conflictos (dos estimadores corrigieron lo mismo): resuelve el importador por
  `modificado_en` con log, mismo criterio que `cloudsync_backend.py`.

### 7.4 Fases (cada una vale por sí sola)

1. **Fase 1 — Workbench de revisión en la app** (sin write-back): reemplaza la
   hoja Revision por UI; exporta PDF + Excel. Rama: `feature/workbench-revision`.
2. **Fase 2 — JSON de propuesta** (contrato v1.3): la sesión de revisión emite
   el payload automático; costeo lo consume con dry-run. Ramas:
   `feature/propuesta-json` (este repo) + `feature/importador-propuestas` (costeo).
3. **Fase 3 — Nube de propuestas**: los JSON aterrizan solos en `propuestas`
   (tokens de contribuidor); importador con cola de curaduría. Rama:
   `feature/nube-propuestas`.

---

## Registro de decisiones que nos llevaron aquí

- `c791aa2` Excel vivo por defecto: el PU nace del SQL de costeo (semilla) y
  una vez en el libro es de Excel (fórmulas encadenadas Insumos → APU →
  Presupuesto). El oficial queda como fila de referencia en cada APU y en
  Revision.
- `d218693` Matcher anti eco-de-subconjunto: nació del caso K78.7 (muebles
  confirmados contra la descripción-cápsula `12"Ø.` de SISTEMA TITE LINER).
- `9d34de6` Líneas `%` de las matrices = % del subtotal MO, no insumos.
- `065ba68` Agregación de insumos por (clave, unidad) — la trampa KG vs PZA/TON.
- `74ecafe` Web: versión del catálogo visible en vivo (manifest del snapshot).
