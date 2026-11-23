# Plan de Implementación — Presupuestador PONTEM

> **Última actualización:** 2026-11-20
> **Estado:** Planificación

---

## Decisión de arquitectura: Plantilla híbrida

| ¿Plantilla externa o generada por código? | **Híbrida** |
|---|---|
| La plantilla `machote.xlsx` contiene lo visual fijo y las hojas FASAR con fórmulas. | El código copia la plantilla y escribe encima los datos dinámicos del proyecto. |
| Se genera **una vez** con un script dedicado y se versiona. | `presupuestar.py` nunca modifica el machote original. |

**Qué va en la plantilla (fijo):**
- Estilos, colores, fuentes, bordes, anchos de columna
- Hoja `Categorias` con estructura y fórmulas FASAR
- Hoja `Parametros` con estructura de cálculo FASAR
- Hoja `Cuadrillas` con estructura de columnas
- Hoja `Calculo_categoria` con desglose por categoría
- Hoja `Notas` con alcance y fundamento

**Qué crea el código en cada ejecución (dinámico):**
- Hoja `Variables` con datos del proyecto (UMA, indirectos, utilidad, IVA, ISN, nombre, cliente, fecha)
- Hoja `Presupuesto` con los conceptos del archivo de entrada y fórmulas `VLOOKUP → Insumos`
- Hoja `Insumos` como tabla maestra consolidada (materiales, MO, equipo, cuadrillas, conceptos_repo)
- Hojas `APU_<clave>` — una por cada concepto desglosado (N dinámico, no fijo en 10)
- Hoja `Revisión` con resultados del matcheo
- Hoja `Explosion` con la explosión de insumos (BOM)

---

## Pasos de implementación

### Paso 1 — Refactorizar plantilla con fórmulas FASAR reales

**Objetivo:** La plantilla actual (`plantilla_presupuesto.py`) tiene valores duros. Hay que reescribir las hojas FASAR con fórmulas nativas de Excel ligadas entre sí, y guardar el resultado como `plantillas/machote.xlsx`.

- [x] Extraer estilos a `src/estilos.py` (colores, fuentes, bordes, alineaciones)
- [x] Reescribir `Categorias` → cada categoría con su `Sn` editable, sin valores duros de FASAR
- [x] Reescribir `Parametros` → estructura completa FASAR con celdas editables (Tp, Tl, UMA, cuotas IMSS, etc.)
- [x] Reescribir `Calculo_categoria` → FASAR y Sr calculados con fórmulas que referencian `Parametros`
- [x] Reescribir `Cuadrillas` → columna `$/Jornal` con fórmula `=CANT1 * VLOOKUP(CAT1, Calculo_categoria, col_Sr, FALSE) + CANT2 * ...`
- [x] Agregar hoja `Variables` base con estructura (UMA, ISN, % indirectos, % utilidad, % IVA, rendimiento)
- [x] Guardar como `plantillas/machote.xlsx`

**Notas / Decisiones:**
- La plantilla vieja (`plantilla_presupuesto.py`) se conserva como referencia. El nuevo generador es `generar_plantilla.py`.
- Las fórmulas FASAR se verificaron correctas: cadena Parametros → Calculo_categoria → Cuadrillas funciona.
- El rango VLOOKUP en Cuadrillas cubre A5:K50 (46 categorías). Si se agregan más, ampliar el rango.
- `Variables` queda con placeholders; `presupuestar.py` los sobrescribirá con los datos reales del proyecto.
- `machote.xlsx` NO incluye hojas `Presupuesto`, `Insumos`, `APU_*`, `Revisión` ni `Explosion` — esas las crea dinámicamente `presupuestar.py`.


---

### Paso 2 — Fase 0: Estructura del proyecto

**Objetivo:** Crear la estructura de carpetas, archivos de configuración y dependencias.

- [x] Crear estructura de carpetas:
  ```
  presupuestador/
  ├─ presupuestar.py
  ├─ config.yaml
  ├─ requirements.txt
  ├─ .gitignore
  ├─ venv/                  ← Entorno virtual (no se versiona)
  ├─ src/
  │  ├─ __init__.py
  │  ├─ estilos.py
  │  ├─ indexador.py
  │  ├─ lector_entrada.py
  │  ├─ lector_fasar.py
  │  ├─ matcher.py
  │  ├─ constructor.py
  │  ├─ formulas.py
  │  ├─ validador.py
  │  └─ llm.py
  ├─ plantillas/
  │  └─ machote.xlsx
  ├─ salidas/
  └─ tests/
  ```
- [x] Crear `requirements.txt` con dependencias (6 activas + 3 comentadas para fases futuras)
- [x] Crear `config.yaml` con defaults: matcher, FASAR, salida, validación
- [x] Verificar `pip install -r requirements.txt` sin errores (21 paquetes instalados en venv)

**Notas / Decisiones:**


---

### Paso 3 — Fase 1: Lectura de datos

**Objetivo:** Poder leer el archivo de entrada y el catálogo del repo PONTEM.

- [x] `src/indexador.py`:
  - Cargar `catalogo_costos.json` (~12,329 conceptos)
  - Índice en memoria por `clave`, `descripcion`, `nivel_1`, `nivel_2`
  - Caché opcional en `pickle` para acelerar re-ejecuciones
- [x] `src/lector_entrada.py`:
  - Leer `.xlsx` con columnas: `id | partida | concepto | unidad | cantidad`
  - Validar que las 5 columnas existen
  - Devolver lista de diccionarios

**Notas / Decisiones:**
- `lector_entrada.py` soporta ambos formatos (4 y 5 columnas). Si no hay `partida`, asigna "GENERAL".
- Se agregó `generar_entrada_ejemplo()` para crear archivos de prueba rápidamente.
- El indexador tiene soporte para caché pickle, pero por ahora cargamos directo del JSON.


---

### Paso 4 — Fase 3: Matcheo

**Objetivo:** Emparejar cada concepto de entrada con su correspondiente en el catálogo PONTEM.

- [x] `src/matcher.py`:
  - Match exacto por `descripcion` normalizada (lowercase, sin acentos, sin espacios extra)
  - Match fuzzy con `rapidfuzz` (umbral configurable, default 85)
  - Para cada concepto, devolver: `clave_sugerida`, `descripcion_sugerida`, `score`, `metodo` (exacto/fuzzy/sin_match)
- [ ] Escribir resultados en hoja `Revisión`:
  - Columnas: `Id entrada | Concepto entrada | Unidad entrada | Cantidad | Clave sugerida | Descripción sugerida | Score | Método | Estado | Desglosar APU`
  - Estados: `confirmado`, `dudoso`, `sin_match`

**Notas / Decisiones:**


---

### Paso 5 — Fase 4-6: Constructor

**Objetivo:** Construir el `.xlsx` de salida completo a partir de la plantilla, los datos de entrada y los resultados del matcheo.

- [x] `src/formulas.py`:
  - Helpers para generar fórmulas Excel: `VLOOKUP`, `SUMIF`, `SUMPRODUCT`, referencias entre hojas
  - Helper para traducir número de columna a letra (1→A, 2→B, ...)
  - Helper para rangos seguros con nombres de hoja (escapar espacios, apóstrofes)
- [x] `src/constructor.py` — Parte 1: Hojas base:
  - Copiar `machote.xlsx` a `salidas/<timestamp>/presupuesto.xlsx`
  - Llenar hoja `Variables` con datos del proyecto + parámetros FASAR
  - Vincular `Categorias` y `Calculo_categoria` a `Variables`
  - Llenar `Cuadrillas` con las cuadrillas necesarias y fórmulas `$/Jornal`
- [x] `src/constructor.py` — Parte 2: Insumos:
  - Consolidar todos los insumos de los APUs en la hoja `Insumos`
  - `tipo=cuadrilla` → fórmula a `Cuadrillas`
  - `tipo=categoria` → fórmula a `Categorias`
  - `tipo=concepto_repo` → valor duro con PU del catálogo
  - `tipo=material/equipo/herramienta` → valor editable
- [x] `src/constructor.py` — Parte 3: APUs dinámicos:
  - Por cada concepto con `Desglosar APU = S` en `Revisión`, crear hoja `APU_<clave>`
  - Estructura: Insumo | Descripción | Unidad | Cantidad | Desperdicio | Cant total | Costo unitario | Importe
  - Fórmulas `VLOOKUP` hacia `Insumos`
  - Subtotal por tipo, costo directo, indirectos, utilidad, PU calculado
- [x] `src/constructor.py` — Parte 4: Presupuesto:
  - Llenar hoja `Presupuesto` con columnas: Partida | Clave | Descripción | Unidad | Cantidad | PU | Importe
  - `Descripción`, `Unidad`, `PU` con `=VLOOKUP(clave, Insumos, col, FALSE)`
  - `Importe = Cantidad × PU`
  - Subtotal por partida con `SUMIF`
  - Totales al pie: CD, indirectos, utilidad, IVA, total general

**Notas / Decisiones:**


---

### Paso 6 — Fase 9: CLI

**Objetivo:** Interfaz de línea de comandos interactiva que orquesta todo el flujo.

- [x] `presupuestar.py`:
  - Usar `typer` para CLI + `rich` para output formateado
  - Preguntar interactivamente:
    - Ruta del archivo de entrada (`.xlsx`)
    - Ruta del repo PONTEM (`catalogo_costos.json`)
    - Datos del proyecto: nombre, cliente, fecha
    - Variables: % indirectos, % utilidad, % IVA
    - Umbral fuzzy `[85]`
  - Mostrar progreso con pasos numerados
  - Mostrar resultado del matcheo en tabla Rich con colores por estado
  - Pedir confirmación antes de construir
  - Si el usuario dice "no", salir para editar `Revisión` y reintentar
  - Validar y guardar

**Notas / Decisiones:**


---

### Paso 7 — Explosión de insumos (BOM)

**Objetivo:** Generar la hoja `Explosion` con la lista consolidada de todos los insumos necesarios para la obra completa.

- [ ] Hoja `Explosion`:
  - Por cada concepto del presupuesto, tomar su cantidad de obra
  - Por cada insumo en su APU, multiplicar: `cantidad_obra × consumo_unitario`
  - Sumar el mismo insumo a través de todos los conceptos
  - Columnas: `Clave insumo | Descripción | Tipo | Unidad | Cantidad total | Costo unitario | Importe total`
  - Agrupar por tipo: materiales, mano de obra, equipo, maquinaria
  - Verificar: `SUMA(importes explosión) = costo directo total del presupuesto`

**Notas / Decisiones:**


---

## Apéndice: Fórmula FASAR (referencia)

```
Fsr = Ps × (Tp / Tl) + (Tp / Tl)

Donde:
  Tp = días pagados al año = 365 + aguinaldo + prima vacacional
  Tl = días efectivamente laborados
  Ps = obligaciones patronales (IMSS, SAR, Infonavit, ISN) como fracción del SBC
  
  Sr = Sn × Fsr
```

**Artículos aplicables:** RLOPSRM 190, 191, 192.

---

## Apéndice: Explosión de insumos (referencia)

```
Para cada concepto del presupuesto:
  cantidad_obra × consumo_unitario_insumo = cantidad_insumo_en_concepto

Sumar el mismo insumo a través de todos los conceptos donde aparece.

Resultado: lista de insumos con cantidad total y unidad.
El importe se obtiene multiplicando cantidad_total × costo_unitario.
```

---

## Bitácora de cambios

| Fecha | Cambio |
|---|---|
| 2026-11-20 | Creación del plan. Definida arquitectura híbrida (plantilla externa + llenado dinámico). |
| 2026-11-20 | **Paso 1 completado.** Creado `src/estilos.py` y `generar_plantilla.py`. Plantilla `machote.xlsx` generada con 6 hojas y fórmulas FASAR nativas verificadas. |
| 2026-11-20 | **Paso 2 completado.** Venv creado, `requirements.txt` (6 dependencias), `config.yaml`, `.gitignore`. `pip install -r` verificado. |
| 2026-11-20 | **Paso 3 completado.** `indexador.py` (carga JSON + búsqueda por clave/texto + caché pickle) y `lector_entrada.py` (soporta 4 y 5 columnas, validación, generador de ejemplo). Test integrado OK. |
| 2026-11-20 | **Paso 4 completado.** `matcher.py` con normalización Unicode, match exacto + fuzzy (rapidfuzz token_sort_ratio), umbral configurable, top-N candidatos. Test: 2/2 exactos + 1 dudoso OK. |
| 2026-11-20 | **Paso 5 completado.** `formulas.py` (VLOOKUP, SUMIF, SUMPRODUCT, refs entre hojas) + `constructor.py` (copia plantilla, escribe Variables/Insumos/Revisión/APU/Presupuesto/Explosión). Test end-to-end: 10 conceptos, 100% match exacto, VLOOKUP funcionando. |
| 2026-11-20 | **Paso 6 completado.** `presupuestar.py` — CLI con typer + rich. Flujo completo: prompts interactivos, indexado, matcheo, tabla de revisión coloreada, confirmación, construcción, guardado. Testeado con entrada simulada. |
| 2026-11-20 | **Paso 7 completado.** Hoja `Explosion` integrada en constructor: agrupa insumos por clave, suma cantidades, importe = cantidad × PU, total cuadra con costo directo. |
| 2026-11-21 | **Soporte de Partidas y Subpartidas.** `lector_entrada.py` detecta filas `PARTIDA`/`SUBPARTIDA` incrustadas (formato de 4 columnas tipo VENADO.xlsx). `constructor.py` escribe encabezados de partida, subtotales por partida, y anidación de 2 niveles (subpartida → partida → general). El gran total suma los subtotales de cada partida. Matcher propaga `subpartida`. Probado con VENADO.xlsx: 217 conceptos, 21 partidas. |
| 2026-11-21 | **Formato tipo mockup OMA + exportación PDF.** `escribir_presupuesto` ahora replica el formato del mockup (Pretil OMA MTY): banda de título gris, nombre de proyecto en verde, cliente/fecha/cotización/revisión, tabla WBS (No./Clave/Descripción/Cantidad/Unidad/PU/Importe/%), partidas azul `0070C0`, subpartidas azul claro, % del total. Nuevo módulo `src/pdf_export.py` exporta a PDF solo `Presupuesto` + cada `APU*` (una por página) vía Excel COM + pywin32 + pypdf. Botón de descarga PDF en Streamlit y generación automática en CLI. Bug de IVA/indirectos corregido (referencias B18 y CIC+CIO). |
| 2026-11-21 | **Fix crítico de inflado (72M → 2.9M).** El matcher pegaba 11 insumos de gas (tubería/codos/válvulas en ml/pza/kg) al concepto agregado "Sistema de aire acondicionado" ($687,506) porque una consulta corta es subconjunto de una descripción gigante (`token_set_ratio=100`). Se añadió: (1) rechazo duro por unidad incompatible (ml vs pza, kg vs m2…); (2) penalización por "concepto agregado" (factor por longitud relativa + cobertura inversa). Resultado VENADO: costo directo $2,912,312; gran total $4,341,768; sensato para una vivienda. |
| 2026-11-21 | **IA opcional para dudosos (DeepSeek/OpenRouter).** Nuevo `src/llm.py` agnóstico (API OpenAI-compatible): prompt estructurado + respuesta JSON, valida que la clave exista entre los candidatos (nunca inventa), aplica a `Revisión` con columnas `Clave IA / Confianza IA / Motivo IA` y estado `confirmado_ia` si `auto_aplicar`. Config en `config.yaml` (sección `llm`) y `.env.example`. Se agregó `requests`. |
| 2026-11-21 | **APUs e Insumos desde las MATRICES del repo. Revisión fuera del Excel/PDF.** Nuevo `src/apu_source.py`: localiza las matrices vía `rutas_costos.py` (`G:\Mi unidad\Costos\PU2026\*\MATRICES\**\*.xlsx`) y extrae 9,348 APUs (clave, unidad, PU, insumos con materiales/MO/equipo) con caché `cache/apu_index.json`. Constructor reescrito: `Insumos` = lista de materiales agregada (Σ cantidad_APU × cantidad_obra); una hoja `APU_<clave>` por concepto con sus insumos reales (VLOOKUP al costo de `Insumos`, recálculo vivo); `Presupuesto` referencia el PU de la hoja APU. Se eliminan `Revisión` y `Explosión` del libro (visto bueno se da en CLI/app, no en el archivo). VENADO: 72 hojas (64 APU + Insumos + fijas), sin Revisión/Explosión, PDF 83 páginas, gran total $3,990,740. |