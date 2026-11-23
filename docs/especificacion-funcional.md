# ESPECIFICACIÓN FUNCIONAL
## Sistema de Presupuestación PONTEM — `presupuestar.py`

---

### 0. Objetivo

Un script Python que construye un presupuesto de construcción en un único `.xlsx` a partir de:

- un archivo de entrada con los conceptos a cotizar,
- el catálogo de precios unitarios del repo PONTEM,
- el archivo de parámetros FASAR,
- una plantilla `.xlsx` de presentación.

El `.xlsx` resultante contiene fórmulas nativas de Excel, de modo que si el usuario modifica un insumo, una cuadrilla o una variable, se recalculan automáticamente los APUs y el presupuesto.

---

### 1. Alcance

El script:

1. Pide por CLI la ruta del archivo de entrada, la ruta del repo PONTEM, la ruta de la plantilla y la ruta del archivo FASAR.
2. Copia la plantilla a una subcarpeta de salida con timestamp. Nunca modifica el archivo original.
3. Indexa el catálogo `catalogo_costos.json` del repo PONTEM.
4. Lee el archivo de entrada con columnas exactas `id`, `concepto`, `unidad`, `cantidad`.
5. Lee el archivo FASAR: parámetros, categorías y cuadrillas.
6. Empareja cada concepto de entrada con un concepto del catálogo PONTEM.
7. Escribe en la copia de la plantilla las hojas: `Presupuesto`, `Variables`, `Insumos`, `Categorías`, `Cuadrillas`, `APU_<clave>` y `Revisión`.
8. Incrusta la calculadora FASAR dentro de la plantilla como hojas nativas (`Variables`, `Categorías`, `Cuadrillas`). El archivo FASAR externo solo se usa durante la construcción.
9. Valida el `.xlsx` construido.
10. Guarda un único archivo de salida.

---

### 2. Contexto del repo PONTEM

El repo PONTEM es un monorepo con carriles de costo. El carril activo es `costos_unitarios/` con ~12,329 conceptos. El índice canónico es `costos_unitarios/catalogo_costos.json`, generado por un pipeline de 4 pasos (consolidador → deduplicador → unificador de PU → generador de índice).

Cada concepto del catálogo tiene las columnas:

```
id, paquete, nivel_1, nivel_2, clave, descripcion, unidad, precio_unitario, aplicacion
```

El repo guarda el precio unitario cerrado de cada concepto, pero no guarda la matriz de insumos que lo compone ni las cuadrillas asociadas. Por esa razón, el presupuesto trata cada concepto del repo como un insumo de tipo `concepto_repo` con cantidad 1 y costo igual a su precio unitario. Los conceptos que el usuario decida desglosar se construyen como hojas `APU_<clave>` con su propia matriz de insumos, ligada a la hoja maestra `Insumos`.

> **Nota:** Para consumir el catálogo de costos desde otros repos o sistemas (sin pasar por el presupuestador), consulta [`docs/consumo-costos.md`](../docs/consumo-costos.md) — 5 niveles de integración, desde lectura directa del JSON hasta sincronización automática con Task Scheduler.

---

### 3. Entradas del sistema

**Archivo de entrada (`.xlsx`)**
Cuatro columnas exactas, en este orden: `id`, `concepto`, `unidad`, `cantidad`. Una fila por concepto a cotizar.

**Repo PONTEM**
Ruta a la carpeta que contiene `costos_unitarios/catalogo_costos.json`. El script no asume ninguna ruta por defecto; el usuario la proporciona en cada ejecución.

**Plantilla (`.xlsx`)**
Archivo de presentación tipo "catálogo muro". Define estilos, layout, portada y formato de partidas. Se copia a la salida y se llena sobre la copia.

**Archivo FASAR (`.xlsx`)**
`Calculadora_FASAR_Mexico_2026_corregida.xlsx`. Aporta parámetros, categorías y cuadrillas. Sus valores se incrustan en la plantilla como hojas nativas.

---

### 4. Salida del sistema

Un único archivo `.xlsx` en `salidas/<timestamp>/presupuesto.xlsx`, con las hojas:

1. `Presupuesto`
2. `Variables`
3. `Insumos`
4. `Categorías`
5. `Cuadrillas`
6. `APU_<clave>` (una por cada concepto desglosado)
7. `Revisión`

Todas las celdas de costo son fórmulas nativas de Excel. Las únicas celdas con valor duro permitido son los precios unitarios del catálogo PONTEM para los insumos de tipo `concepto_repo`, porque así vienen del repo y no tienen desglose.

---

### 5. Flujo de ejecución

```
$ python presupuestar.py

? Archivo de entrada (.xlsx): ./entrada/conceptos.xlsx
? Ruta del repo PONTEM:       ./pontem-costos
? Plantilla (machote .xlsx):  ./plantillas/machote.xlsx
? Archivo FASAR (.xlsx):      ./parametros/Calculadora_FASAR_Mexico_2026_corregida.xlsx
? ¿Usar LLM para mapeos dudosos? [s/N]: s
? Umbral fuzzy (0-100) [85]:

[1/6] Indexando repo PONTEM...          ✓ 12,329 conceptos
[2/6] Leyendo entrada...                ✓ 48 conceptos
[3/6] Leyendo FASAR...                  ✓ 3 categorías, 9 cuadrillas
[4/6] Matcheando...                     ✓ 41 exactos, 5 fuzzy, 2 dudosos
[5/6] Escribiendo hoja Revisión...      ✓ 48 filas
? ¿Confirmar mapeos y construir? [S/n]: S
[6/6] Construyendo presupuesto...       ✓ fórmulas escritas
[✓] Validación: 0 errores, 0 advertencias
[✓] Salida: ./salidas/2026-11-19_1430/presupuesto.xlsx
```

Si el usuario responde "no" a confirmar, el script abre la hoja `Revisión` para edición y espera Enter para releer y continuar.

---

### 6. Hoja `Presupuesto`

Formato de partidas con conceptos. Columnas:

| Col | Campo | Tipo |
|---|---|---|
| A | Partida | texto |
| B | Clave | texto |
| C | Descripción | fórmula `=VLOOKUP(B, Insumos, col_desc, FALSE)` |
| D | Unidad | fórmula `=VLOOKUP(B, Insumos, col_unidad, FALSE)` |
| E | Cantidad | número |
| F | PU | fórmula `=VLOOKUP(B, Insumos, col_costo, FALSE)` |
| G | Importe | fórmula `=E*F` |
| H | Desglose | texto: `PU cerrado` o `APU_<clave>` |
| I | Subtotal partida | fórmula `=SUMIF(A, partida, G)` |
| J | % Indirectos | fórmula `=Variables!$B$indirectos` |
| K | % Utilidad | fórmula `=Variables!$B$utilidad` |
| L | Total partida | fórmula `=I*(1+J+K)` |

Al pie: subtotal, indirectos, utilidad, IVA y total general, todos como fórmulas.

---

### 7. Hoja `Variables`

Parámetros globales editables. El ISN se copia tal cual del archivo FASAR y no se modifica. Los demás valores quedan como celdas editables.

| Variable | Valor |
|---|---|
| UMA diaria | del FASAR |
| Tope SBC (UMA) | del FASAR |
| Días de trabajo por mes | 30 |
| % Indirectos | editable |
| % Utilidad | editable |
| % IVA | 16% |
| ISN | del FASAR, fijo |
| Rendimiento default | 1 |

---

### 8. Hoja `Insumos`

Tabla maestra única de todos los insumos: materiales, equipo, herramienta, cuadrillas, categorías y conceptos del repo.

| Col | Campo | Tipo |
|---|---|---|
| A | Clave | texto único |
| B | Tipo | `material`, `equipo`, `herramienta`, `cuadrilla`, `categoria`, `concepto_repo` |
| C | Descripción | texto |
| D | Unidad | texto |
| E | Costo base | número o fórmula |
| F | Fecha precio | fecha |
| G | Proveedor o fuente | texto |
| H | Paquete / nivel_1 / nivel_2 | del repo |
| I | PU repo (solo `concepto_repo`) | número |
| J | Notas | texto |

Reglas de la columna E:
- `tipo=cuadrilla`: fórmula que consulta la hoja `Cuadrillas`.
- `tipo=categoria`: fórmula que consulta la hoja `Categorías`.
- `tipo=concepto_repo`: valor duro con el PU del catálogo PONTEM.
- `tipo=material`, `equipo`, `herramienta`: valor editable por el usuario.

---

### 9. Hoja `Categorías`

Réplica de la hoja `Calculo_categoria` del archivo FASAR, incrustada en la plantilla. Columnas: Categoría, Sueldo, FASAR, Sr, SBC, E&M cuota fija, E&M excedente 3 UMA, E&M prestaciones, E&M gastos médicos, Invalidez y vida, Guarderías, Riesgo de trabajo, SAR / Retiro, Cesantía y vejez, Infonavit, ISN, Total cargas diarias, $/jornal.

Los parámetros de entrada de estas fórmulas se leen de la hoja `Variables`.

---

### 10. Hoja `Cuadrillas`

Composición de cada cuadrilla y su costo por jornada.

| Cuadrilla | Cant. 1 | Categoría 1 | Cant. 2 | Categoría 2 | Cant. 3 | Categoría 3 | $/Jornal |
|---|---|---|---|---|---|---|---|
| MO001 | 0.1 | Cabo de oficios | 2 | Ayudante General | | | fórmula |

La columna `$/Jornal` es una fórmula que suma, para cada renglón, la cantidad por el $/jornal de la categoría correspondiente en la hoja `Categorías`. Agregar una cuadrilla nueva consiste en agregar una fila con la misma fórmula.

---

### 11. Hoja `APU_<clave>`

Una hoja por cada concepto que el usuario marque como "desglosar" en la hoja `Revisión`. Estructura:

| Insumo (clave) | Descripción | Unidad | Cantidad | Desperdicio | Cantidad total | Costo unitario | Importe |
|---|---|---|---|---|---|---|---|
| MAT-001 | Cemento gris 50 kg | saco | 0.5 | 5% | `=D*(1+E)` | `=VLOOKUP(A, Insumos, costo, FALSE)` | `=F*G` |

Al pie: subtotal por tipo (material, mano de obra, equipo, herramienta), costo directo, indirectos, utilidad y PU calculado. La hoja `Presupuesto` referencia este PU mediante `=APU_<clave>!$H$total`.

---

### 12. Hoja `Revisión`

Siempre presente, con o sin LLM.

| Id entrada | Concepto entrada | Unidad entrada | Cantidad | Clave sugerida | Descripción sugerida | Score | Método | Estado | Desglosar APU |
|---|---|---|---|---|---|---|---|---|---|

Estados posibles: `confirmado`, `dudoso`, `sin_match`. La columna `Desglosar APU` la llena el usuario con S o N.

---

### 13. Dependencias entre hojas

```
Variables ──► Categorías ──► Cuadrillas ──┐
                                          ├──► Insumos ──► APU_<clave> ──► Presupuesto
Insumos (materiales, equipo, herramienta) ┘
Repo PONTEM ──► Insumos (tipo=concepto_repo) ──► Presupuesto
```

Reglas:
- `Presupuesto` nunca lee directamente del repo; siempre pasa por `Insumos`.
- `APU_<clave>` solo lee `Insumos`.
- `Cuadrillas` solo lee `Categorías`.
- `Categorías` solo lee `Variables` y sus propios valores FASAR.
- Cambiar un material en `Insumos` recalcula todos los APUs que lo usan y, por tanto, los PU en `Presupuesto`.

---

### 14. Módulos

| Módulo | Entrada | Salida |
|---|---|---|
| `indexador.py` | `catalogo_costos.json` | índice en memoria y caché opcional |
| `lector_entrada.py` | `.xlsx` de entrada | lista de conceptos (`id`, `concepto`, `unidad`, `cantidad`) |
| `lector_fasar.py` | `.xlsx` FASAR | parámetros, categorías y cuadrillas |
| `matcher.py` | conceptos de entrada e índice | mapeo con score y método |
| `llm.py` | candidatos dudosos | top-3 sugerencias con justificación |
| `constructor.py` | mapeo confirmado, plantilla y FASAR | copia de la plantilla con hojas y fórmulas |
| `formulas.py` | — | helpers para escribir fórmulas |
| `validador.py` | `.xlsx` construido | reporte de errores |

---

### 15. Política del LLM

- Se invoca solo cuando el score fuzzy es menor al umbral configurado y el usuario lo activó.
- Recibe la descripción de entrada y los 10 mejores candidatos del repo (clave, descripción, unidad, PU).
- Devuelve un ranking con justificación corta. No inventa claves.
- Nunca escribe en `Presupuesto`, `Insumos` ni `APU_<clave>`. Solo escribe en `Revisión`.
- Siempre hay confirmación humana antes de construir.
- Si el usuario no activa el LLM, el flujo funciona igual y los casos dudosos quedan como `sin_match` para resolverse a mano en `Revisión`.

---

### 16. Validaciones antes de guardar

1. Toda clave en `Presupuesto` existe en `Insumos`.
2. Toda clave en `APU_<clave>` existe en `Insumos`.
3. Toda categoría en `Cuadrillas` existe en `Categorías`.
4. Las unidades entre `Presupuesto` e `Insumos` son consistentes.
5. No hay ciclos de referencia entre hojas.
6. Las fórmulas tienen sintaxis válida.
7. Los totales cuadran.

Si hay errores, el archivo se guarda como `presupuesto_con_errores.xlsx` para depuración y no como entrega final.

---

### 17. Dependencias Python

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

### 18. Plan de implementación por fases

| Fase | Entregable | Criterio de aceptación |
|---|---|---|
| F0 | Estructura de carpetas, `config.yaml`, `requirements.txt` | `pip install -r` sin errores |
| F1 | `indexador.py` y `lector_entrada.py` | Lee los 12,329 conceptos y el archivo de entrada |
| F2 | `lector_fasar.py` | Extrae parámetros, categorías y cuadrillas correctamente |
| F3 | `matcher.py` sin LLM | Al menos 90% de la entrada mapeada o marcada |
| F4 | `constructor.py` con `Variables`, `Categorías`, `Cuadrillas`, `Insumos` incrustadas | Abre en Excel o LibreOffice sin errores |
| F5 | `constructor.py` con `Presupuesto` y `Revisión` | Fórmulas `VLOOKUP` funcionan |
| F6 | `constructor.py` con `APU_<clave>` | Cambiar un insumo mueve el PU |
| F7 | `validador.py` | Reporte limpio en caso feliz y detección de errores inyectados |
| F8 | `llm.py` | Sugerencias en `Revisión` sin tocar el resto |
| F9 | CLI final y pruebas end-to-end | Corre con los tres `.xlsx` de ejemplo del usuario |

---

### 19. Estructura de carpetas

```
presupuestador/
├─ presupuestar.py
├─ config.yaml
├─ requirements.txt
├─ src/
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
│  └─ 2026-11-19_1430/
│     └─ presupuesto.xlsx
└─ tests/
```

El script siempre copia `plantillas/machote.xlsx` a `salidas/<timestamp>/presupuesto.xlsx` y trabaja sobre la copia.