# Sistema de Presupuestación PONTEM

Sistema automatizado para construir presupuestos de construcción en Excel a partir del catálogo de precios unitarios del repositorio PONTEM. Genera un archivo `.xlsx` con fórmulas nativas, de modo que cualquier modificación en insumos, cuadrillas o variables recalcula automáticamente los Análisis de Precios Unitarios (APU) y el presupuesto completo.

---

## Objetivo

Un script Python (`presupuestar.py`) que construye un presupuesto de construcción en un único archivo `.xlsx` a partir de:

- Un archivo de entrada con los conceptos a cotizar.
- El catálogo de precios unitarios del repo PONTEM (~12,329 conceptos).
- El archivo de parámetros FASAR (categorías, cuadrillas y variables).
- Una plantilla `.xlsx` de presentación.

El resultado es un libro de Excel autocontenido, con dependencias encadenadas en fórmulas nativas, sin necesidad de ejecutar el script nuevamente al ajustar valores.

---

## Alcance funcional

1. **CLI interactiva**: solicita las rutas del archivo de entrada, repo PONTEM, plantilla y archivo FASAR.
2. **Copia segura de la plantilla**: nunca modifica el archivo original; trabaja sobre una copia con timestamp en `salidas/`.
3. **Indexación del catálogo**: carga e indexa `costos_unitarios/catalogo_costos.json`.
4. **Lectura de entrada**: procesa un `.xlsx` con columnas `id`, `concepto`, `unidad`, `cantidad`.
5. **Lectura FASAR**: extrae parámetros, categorías y cuadrillas del archivo de cálculo FASAR.
6. **Emparejamiento automático**: matchea cada concepto de entrada contra el catálogo PONTEM (exacto, fuzzy y opcionalmente con LLM).
7. **Construcción del libro**: escribe las hojas `Presupuesto`, `Variables`, `Insumos`, `Categorías`, `Cuadrillas`, `APU_<clave>` y `Revisión`.
8. **Validación**: revisa integridad de fórmulas, referencias y totales antes de la entrega.
9. **Archivo único de salida**: `salidas/<timestamp>/presupuesto.xlsx`.

---

## Hojas del libro generado

| Hoja | Contenido |
|---|---|
| **Presupuesto** | Partidas con conceptos, cantidades, PU e importes. Fórmulas `VLOOKUP` hacia `Insumos` y `APU_<clave>`. Totales con indirectos, utilidad e IVA. |
| **Variables** | Parámetros globales editables: UMA, indirectos, utilidad, IVA, ISN, días de trabajo, rendimiento default. |
| **Insumos** | Tabla maestra única: materiales, equipo, herramienta, cuadrillas, categorías y conceptos del repo. Columna de costo como valor editable o fórmula según el tipo. |
| **Categorías** | Cálculo FASAR de costos por categoría de mano de obra, ligado a `Variables`. |
| **Cuadrillas** | Composición de cuadrillas (categorías + cantidades) y su $/jornal calculado con fórmulas. |
| **APU_\<clave\>** | Desglose completo de cada concepto seleccionado: insumos, cantidades, desperdicio, importe y costo directo. |
| **Revisión** | Resultado del matcheo: clave sugerida, score, método, estado y columna para marcar desglose de APU. |

### Dependencias entre hojas

```
Variables ──► Categorías ──► Cuadrillas ──┐
                                          ├──► Insumos ──► APU_<id> ──► Presupuesto
Insumos (materiales, equipo, herramienta) ┘
Repo PONTEM ──► Insumos (tipo=concepto_repo) ──► Presupuesto
```

### IDs locales (liga dentro del libro)

Todo lo que se copia de la base general (catálogo/matrices) recibe un **ID local**
del libro y **todas las referencias lo usan**:

| Hoja | Columna | Contenido |
|---|---|---|
| **Insumos** | A `Id local` | `INS-001`, `INS-002`… generado para esta obra |
| **Insumos** | I `Clave origen` | clave de la base general (trazabilidad) |
| **APU_\<id\>** | A `Id local` | el mismo `INS-xxx` que en Insumos (liga por VLOOKUP) |
| **APU_\<id\>** | G `Clave origen` | clave de la base general |
| **Cuadrillas** | A `Id local` | el mismo `INS-xxx`, para que el VLOOKUP del $/Jornal resuelva |
| **Presupuesto** | A `Id` | id local del concepto (ej. `CIM-01`), que es el nombre de su hoja APU |

Con esto el libro es **autocontenido**: ninguna fórmula depende de claves que
existen solo en la base general. Las cuadrillas copiadas se agregan a la hoja
`Cuadrillas` con su $/Jornal de origen (columna I) para que la mano de obra no
quede en `$0`.

---

## 🔌 Integración con otros repos

El catálogo de costos PONTEM está diseñado para ser consumido desde cualquier otro repositorio o sistema. Hay cinco niveles de integración, ordenados de más simple a más integrada:

| Nivel | Método | Ideal para |
|---|---|---|
| **1** | Leer `catalogo_costos.json` directamente | Cualquier lenguaje (JS, Go, Rust); cero dependencias |
| **2** | `pip install` del paquete PONTEM | API limpia en Python con scoring: `buscar("muro de block")` |
| **3** | Copiar `buscador_costos.py` (~300 líneas, stdlib-only) | Portabilidad máxima, sin instalar nada ni acceder a GitHub |
| **4** | Usar `indice_pu.json` por proyecto | Navegar archivos Excel de un proyecto específico sin abrirlos |
| **5** | Script de sincronización + Task Scheduler | Integración desacoplada con actualización semanal automática |

→ Detalle completo en [`docs/consumo-costos.md`](docs/consumo-costos.md).

---

## Requisitos

- Python 3.10+
- Dependencias (ver `requirements.txt`):

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

## Estructura del proyecto

```
Presupuestos-excel/
├─ presupuestar.py          # CLI (línea de comandos)
├─ app.py                    # Interfaz web (Streamlit)
├─ generar_plantilla.py      # Genera plantillas/machote.xlsx
├─ instalar.bat              # Instalación inicial (venv + deps + plantilla)
├─ iniciar_streamlit.bat     # Abre la interfaz web
├─ presupuestar.bat          # Abre la CLI
├─ config.yaml               # Configuración general
├─ requirements.txt          # Dependencias Python
├─ src/
│  ├─ estilos.py             # Colores, fuentes, bordes
│  ├─ visual_dna.py          # Tokens visuales (DNA) de la marca
│  ├─ indexador.py           # Índice sobre costos.db (SQLite)
│  ├─ lector_entrada.py      # Lee el .xlsx de entrada (partidas/subpartidas)
│  ├─ costeo_client.py       # Cliente del repo de costos (Nivel 1 / Nivel 3)
│  ├─ matcher.py             # Emparejamiento exacto / oficial / fuzzy
│  ├─ llm.py                 # Sugerencias con IA (DeepSeek / OpenRouter)
│  ├─ apu_db.py              # APUs desde costos.db
│  ├─ apu_source.py          # APUs desde las matrices .xlsx (con caché)
│  ├─ constructor.py         # Construye el .xlsx (Presupuesto, APU, Portada...)
│  ├─ validador.py           # Validación de integridad del libro generado
│  ├─ formulas.py            # Helpers para fórmulas Excel
│  └─ pdf_export.py          # Exporta Presupuesto + APUs a PDF
├─ plantillas/
│  └─ machote.xlsx           # Plantilla base (FASAR + Variables)
├─ salidas/                  # Salidas generadas (desde la app web)
└─ tests/                    # Pruebas y datos de ejemplo
```

---

## Instalación (una sola vez)

```bat
python -m venv venv
venv\Scripts\Activate.ps1
pip install -r requirements.txt
python generar_plantilla.py
```

> En PowerShell, si bloquea la activación del entorno, ejecuta antes:
> ```powershell
> Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
> ```

---

## Comandos

### Interfaz web (Streamlit)

```bat
venv\Scripts\Activate.ps1
streamlit run app.py
```

Se abre en `http://localhost:8501`. Sube el `.xlsx`, llena los datos y descarga el `.xlsx` y el `.pdf`.

O simplemente haz doble clic en **`iniciar_streamlit.bat`**.

### Línea de comandos (CLI)

```bat
venv\Scripts\Activate.ps1
python presupuestar.py
```

También acepta la ruta del archivo como argumento:

```bat
python presupuestar.py "C:\ruta\a\mi_archivo.xlsx"
```

Y puedes usar **`presupuestar.bat`** con doble clic o pasando la ruta:

```bat
presupuestar.bat "C:\ruta\a\mi_archivo.xlsx"
```

### Regenerar la plantilla base

```bat
venv\Scripts\Activate.ps1
python generar_plantilla.py
```

### Pruebas

Las pruebas son autocontenidas (no necesitan el repo de costos ni las matrices):

```bat
venv\Scripts\Activate.ps1
python tests\run_all.py
```

O una suite individual:

```bat
python tests\test_matcher.py
python tests\test_constructor.py
python tests\test_validador.py
```

El CI (`.github/workflows/ci.yml`) ejecuta `tests/run_all.py` en cada push.

### Validar un libro ya generado

```bat
python -m src.validador "salidas\presupuesto.xlsx"
```

Devuelve 0 si el libro no tiene errores y 1 si tiene alguno.

---

## Uso

```bash
python presupuestar.py
```

El archivo de entrada `.xlsx` puede tener dos formatos:

**Formato A — columnas con partida explícita** (`id | partida | concepto | unidad | cantidad`)

**Formato B — partidas incrustadas como filas** (`id | concepto | unidad | cantidad`),
donde las filas cuyo `id` empieza con `PARTIDA` o `SUBPARTIDA` agrupan los conceptos siguientes:

```
ID                              | Concepto              | Unidad | Cantidad
PARTIDA 01: CIMENTACIÓN         |                       |        |
CIM-01                          | Excavación en cepa... | m3     | 9.75
CIM-02                          | Relleno y compact...  | m3     | 0.50
SUBPARTIDA 01.1: ZAPATAS        |                       |        |
CIM-03                          | Zapata aislada ZA-1   | pza    | 4.00
PARTIDA 02: MUROS               |                       |        |
MPB-01                          | Castillos de concreto | ml     | 54.00
```

El sistema genera automáticamente en la hoja `Presupuesto`:
- Un **encabezado** por cada partida.
- Los conceptos agrupados bajo su partida.
- Un **subtotal por partida** (y por subpartida, si existen).
- El gran total que suma los subtotales de todas las partidas.

También puedes pasar la ruta como argumento:

```bash
python presupuestar.py "ruta/al/archivo.xlsx"
```

Si pasas una carpeta, busca los `.xlsx` dentro y te deja elegir.

### Cómo se buscan los precios (consumo por contrato)

El presupuestador **consume el repo de costos según el contrato** documentado en [`docs/consumo-costos.md`](docs/consumo-costos.md):

1. **Descubre el repo entre carpetas hermanas** automáticamente, priorizando el que tiene `buscador_costos.py` y `costos_unitarios/catalogo_costos.json` (por ejemplo `C:\Pontem_dev\costeo`). Excluye la propia carpeta de la app.
2. **Nivel 3 — buscador oficial:** si existe `buscador_costos.py`, lo importa y usa su función `buscar()` como fuente de candidatos. Si no, cae a **Nivel 1** (lectura directa del JSON). La app muestra qué nivel usó.
3. Cada concepto de entrada se **re-rankea** con: normalización, stopwords, filtro/penalización por unidad y cobertura ponderada por rareza (IDF).
4. Resultado por concepto:
   - **confirmado** — score ≥ umbral (por defecto 55). Precio tomado del catálogo.
   - **dudoso** — score entre el piso (40) y el umbral. Precio del catálogo pero **revisar**.
   - **sin_match** — score < piso. Se muestra en el presupuesto con **PU $0 y fila roja** para completarlo.
5. Las líneas dudosas con un precio anómalo (>12× la mediana de su unidad) se dejan en **$0** para no inflar el total.
6. **Nada se pierde:** el `Presupuesto` incluye todos los conceptos. Las filas rojas son las que debes completar en la hoja `Revisión`.

> Puedes forzar un repo específico pasando su ruta como `repo_path` (app web) o editando la ruta en la CLI.

### IA para conceptos dudosos (opcional)

El matching automático no siempre acierta en descripciones libres. Con IA, el script manda un **prompt** (concepto + mejores candidatos del catálogo) al modelo, que devuelve **JSON** con la clave que mejor corresponde. No es un chat interactivo: es automático.

**Proveedores** (ambos OpenAI-compatibles):

| Provider | Modelo por defecto | Llave |
|---|---|---|
| `deepseek` (directo) | `deepseek-chat` (o `deepseek-reasoner`) | `DEEPSEEK_API_KEY` |
| `openrouter` | `deepseek/deepseek-chat` (o `deepseek/deepseek-r1`) | `OPENROUTER_API_KEY` |

**Configuración:**

```bat
copy .env.example .env
```

Edita `.env`:

```env
LLM_PROVIDER=deepseek
LLM_MODEL=deepseek-chat
DEEPSEEK_API_KEY=sk-...
```

**Uso:** al ejecutar la CLI o la app, cuando hay conceptos dudosos/sin match aparece la opción de usar IA. Las sugerencias se escriben en la hoja **`Revisión`** (columnas `Clave IA`, `Confianza IA`, `Motivo IA`). Si `auto_aplicar: true` en `config.yaml`, la clave elegida se usa en el presupuesto con estado `confirmado_ia`.

- La IA **solo puede elegir entre los candidatos del catálogo**; nunca inventa claves (se valida antes de aplicar).
- Sin llave configurada, todo funciona igual: los dudosos quedan para revisión manual.
- Parámetros en `config.yaml` → sección `llm` (`lote`, `max_items`, `max_candidatos`, `auto_aplicar`, `temperature`).

### Salida

El script genera en la **misma carpeta del archivo de origen**:

- `presupuesto.xlsx` — libro completo con fórmulas nativas.
- `presupuesto.pdf` — PDF con **solo la hoja Presupuesto y cada APU en su página individual** (ideal para imprimir).

La hoja `Presupuesto` sigue el formato de catálogo tipo OMA: encabezado con nombre del proyecto, cliente, fecha, cotización y revisión; tabla WBS con `No. / Clave / Descripción / Cantidad / Unidad / PU / Importe / %`; partidas en azul con subtotales y porcentaje del total; pie con costo directo, indirectos, utilidad, IVA y gran total.

### App web

```bash
streamlit run app.py
```

Sube el archivo, llena los datos del proyecto y descarga el `.xlsx` y el `.pdf`.

---

## Flujo de ejecución

```
[1/6] Indexando repo PONTEM...          ✓ 12,329 conceptos
[2/6] Leyendo entrada...                ✓ N conceptos
[3/6] Leyendo FASAR...                  ✓ N categorías, N cuadrillas
[4/6] Matcheando...                     ✓ N exactos, N fuzzy, N dudosos
[5/6] Escribiendo hoja Revisión...      ✓ N filas
[6/6] Construyendo presupuesto...       ✓ fórmulas escritas
[✓] Validación: 0 errores, 0 advertencias
[✓] Salida: ./salidas/2026-11-19_1430/presupuesto.xlsx
```

---

## Política del LLM

- Solo se invoca si el usuario lo activa y el score fuzzy está por debajo del umbral.
- Recibe la descripción de entrada y los 10 mejores candidatos del catálogo.
- Devuelve un top-3 con justificación corta; nunca inventa claves ni escribe fuera de `Revisión`.
- Siempre requiere confirmación humana antes de construir.
- Si no se activa, los casos dudosos quedan como `sin_match` para resolverse manualmente.

---

## Validaciones

Al terminar de construir el libro, `src/validador.py` lo revisa y reporta el
resultado en la salida de la CLI. Las verificaciones son:

1. Toda clave usada en un `VLOOKUP` a Insumos/Cuadrillas existe en esa hoja.
2. Toda clave de insumo de las hojas APU existe en Insumos.
2b. Los IDs locales de Insumos son únicos (si no, el VLOOKUP tomaría el costo
   de otro insumo).
3. Toda categoría usada en Cuadrillas existe en `Calculo_categoria`.
4. Consistencia de unidades entre `Presupuesto`, las hojas APU e `Insumos`.
5. Ausencia de ciclos de referencia entre hojas.
6. Sintaxis válida de las fórmulas (paréntesis balanceados, sin `#REF!`, sin
   referencias a hojas inexistentes).
7. Cuadre de totales: se verifica la cadena de fórmulas (CD → indirectos →
   utilidad → IVA → gran total) y, si hay Excel disponible, se recalcula una
   copia del libro y se comparan los importes con la suma de los conceptos.
8. Ninguna cuadrilla queda en `$0` (el `IFERROR` del VLOOKUP lo ocultaría y el
   presupuesto quedaría subvaluado) y la tabla de la Portada suma el Costo
   Directo del Presupuesto.

Si el libro tiene errores, además se guarda una copia como
`presupuesto_con_errores.xlsx` en la misma carpeta para depuración.

---

## Plan de implementación

| Fase | Entregable | Criterio |
|---|---|---|
| F0 | Estructura, `config.yaml`, `requirements.txt` | `pip install -r` sin errores |
| F1 | `indexador.py` y `lector_entrada.py` | Carga los 12,329 conceptos y el archivo de entrada |
| F2 | `lector_fasar.py` | Extrae parámetros, categorías y cuadrillas |
| F3 | `matcher.py` sin LLM | ≥90% de la entrada mapeada o marcada |
| F4 | `constructor.py`: hojas base | Abre en Excel/LibreOffice sin errores |
| F5 | `constructor.py`: `Presupuesto` y `Revisión` | Fórmulas `VLOOKUP` funcionan |
| F6 | `constructor.py`: `APU_<clave>` | Cambiar un insumo recalcula el PU |
| F7 | `validador.py` | Reporte limpio en caso feliz; detección de errores inyectados |
| F8 | `llm.py` | Sugerencias en `Revisión` sin tocar otras hojas |
| F9 | CLI final y pruebas end-to-end | Corre con archivos de ejemplo reales |