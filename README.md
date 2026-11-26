# Sistema de Presupuestación PONTEM

Sistema automatizado para construir presupuestos de construcción en Excel a partir del catálogo de precios unitarios del repositorio PONTEM. Genera un archivo `.xlsx` con fórmulas nativas, de modo que cualquier modificación en insumos, cuadrillas o variables recalcula automáticamente los Análisis de Precios Unitarios (APU) y el presupuesto completo.

---

## Objetivo

Un script Python (`presupuestar.py`) que construye un presupuesto de construcción en un único archivo `.xlsx` a partir de:

- Un archivo de entrada con los conceptos a cotizar.
- El catálogo de precios unitarios del repo PONTEM (12,927 conceptos).
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
7. **Construcción del libro**: escribe las hojas `Presupuesto`, `Variables`, `Insumos`, `Categorías`, `Cuadrillas`, `APU_<id>` y `Revisión`.
8. **Validación**: revisa integridad de fórmulas, referencias y totales antes de la entrega.
9. **Archivo único de salida**: `salidas/<timestamp>/presupuesto.xlsx`.

---

## Hojas del libro generado

| Hoja | Contenido |
|---|---|
| **Presupuesto** | Partidas con conceptos, cantidades, PU e importes. El PU es el **oficial del catálogo** (por defecto) o el de la hoja APU (modo `--cotizar-apu`). Totales con indirectos, utilidad e IVA. |
| **Variables** | Parámetros globales editables: UMA, indirectos, utilidad, IVA, ISN, días de trabajo, rendimiento default. |
| **Insumos** | Tabla maestra única: materiales, equipo, herramienta, cuadrillas, categorías y conceptos del repo. Columna de costo como valor editable o fórmula según el tipo. |
| **Categorías** | Cálculo FASAR de costos por categoría de mano de obra, ligado a `Variables`. |
| **Cuadrillas** | Composición de cuadrillas (categorías + cantidades) y su $/jornal calculado con fórmulas. |
| **APU_\<id\>** | Desglose completo de cada concepto: insumos, cantidades, desperdicio, importe, costo directo, `PU DE VENTA` y `PU OFICIAL (CATALOGO)`. |
| **Revisión** | Resultado del matcheo: clave sugerida, score, método, estado, columna para marcar desglose de APU y la **trazabilidad del precio**: `PU oficial`, `PU APU`, `Dif. %`, `Origen precio`, `Fuente desglose`. |

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

El catálogo de costos PONTEM se consume por su **superficie publicada**
(`dist/contrato_costos/<versión>/`). Hay cinco niveles de integración, de más
simple a más integrado:

| Nivel | Método | Ideal para |
|---|---|---|
| **1** | Leer `catalogo_costos.json` directamente | Cualquier lenguaje (JS, Go, Rust); cero dependencias |
| **2** | `pip install` del paquete PONTEM | API limpia en Python con scoring: `buscar("muro de block")` |
| **3** | Copiar `buscador_costos.py` (stdlib-only) | Portabilidad máxima, sin instalar nada ni acceder a GitHub |
| **4** | La BD canónica (`costos_canonico.db`, read-only) | **APU, rendimientos, cuadrillas, costo horario y el origen del precio** |
| **5** | Script de sincronización + Task Scheduler | Integración desacoplada con actualización semanal automática |

→ Detalle completo en [`docs/consumo-costos.md`](docs/consumo-costos.md).

---

## 📚 Documentación

| Documento | Qué es |
|---|---|
| [`README.md`](README.md) | Guía de uso (este archivo) |
| [`docs/especificacion-funcional.md`](docs/especificacion-funcional.md) | Especificación vigente: hojas, columnas, fórmulas, validaciones |
| [`docs/consumo-costos.md`](docs/consumo-costos.md) | Cómo consumir el contrato de costos desde otros repos |
| [`docs/BACKLOG.md`](docs/BACKLOG.md) | **Registro de seguimiento**: qué está hecho, qué falta, riesgos |
| [`docs/plan.md`](docs/plan.md) | Bitácora histórica del proyecto (no se actualiza) |

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
│  ├─ indexador.py           # Índice de búsqueda sobre el catálogo (JSON/SQLite)
│  ├─ lector_entrada.py      # Lee el .xlsx de entrada (partidas/subpartidas)
│  ├─ costeo_client.py       # Cliente del repo de costos (Nivel 1 / Nivel 3)
│  ├─ matcher.py             # Emparejamiento exacto / oficial / fuzzy
│  ├─ llm.py                 # Sugerencias con IA (DeepSeek / OpenRouter)
│  ├─ apu_db.py              # APUs desde la BD canónica de costos (read-only)
│  ├─ apu_source.py          # Fallback: APUs desde las matrices .xlsx (no contractual)
│  ├─ constructor.py         # Construye el .xlsx (Presupuesto, APU, Portada...)
│  ├─ validador.py           # Validación de integridad del libro generado
│  ├─ formulas.py            # Helpers para fórmulas Excel
│  └─ pdf_export.py          # Exporta Presupuesto + APUs a PDF
├─ plantillas/
│  └─ machote.xlsx           # Plantilla base (FASAR + Variables)
├─ cache/                    # Índice local de matrices (NO se versiona; se regenera)
├─ scripts/
│  └─ diagnostico/           # Scripts de inspección y depuración (no son la suite)
├─ salidas/                  # Salidas generadas (desde la app web)
└─ tests/                    # Suite de pruebas (autocontenida; ver tests/README.md)
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

El presupuestador **consume el repo de costos por su superficie pública**, tal
como define el contrato v1.2 del repo de costeos
([`LEEME.md` del paquete publicado](https://github.com/pontem-dev/costeo) →
*"El consumidor se ata al artefacto, no al interior del repo Costos. No leas
Excel, carpetas fuente ni tablas internas"*).

| Superficie | Para qué la usa el presupuestador | Dónde |
|---|---|---|
| `catalogo_costos.json` | clave → precio unitario (nivel 1) | `costos_unitarios/catalogo_costos.json` o el paquete publicado `dist/contrato_costos/<versión>/` |
| `buscador_costos.py` | búsqueda de candidatos (nivel 3) | raíz del repo de costos |
| `costos_canonico.db` (read-only) | APU, rendimientos, costo horario, origen del precio | `dist/contrato_costos/<versión>/costos_canonico.db` |
| `v_concepto_pu` | **el PU que manda** y su origen | vista pública |
| `v_apu_resumen` / `v_apu_linea_costo` | desglose del APU valuado (cantidad o rendimiento) | vistas públicas |

Reglas que respeta el código:

- La BD canónica se abre **read-only** (`?mode=ro`) y solo se consultan **vistas
  públicas**; nunca tablas internas.
- **No** se leen los Excel de `PU2026/` ni las tablas planas de `costos.db`.
- Se respeta la **prioridad del PU del modelo canónico**:
  `publicado > APU > manual`. Es decir, el PU cotizado es
  `v_concepto_pu.precio_unitario`.
- Si la canónica no trae el APU de un concepto, el **desglose** se completa con
  el índice de matrices (fuente *no contractual*): la hoja Revisión lo marca
  como `matrices` y el PU sigue siendo el oficial. Se desactiva con
  `PONTEM_COSTOS_SIN_MATRICES=1`.

> **El índice de matrices (`cache/apu_index.json`) es local y no se versiona.**
> Es un artefacto derivado de `G:\Mi unidad\Costos\PU2026\` (12 MB), así que
> queda fuera de git (`.gitignore`). Se regenera solo cuando cambian las
> matrices: el código guarda una huella (hash de ruta + fecha + tamaño de cada
> Excel) en `cache/apu_index.meta.json` y rehace el índice si no coincide.
> La primera vez tarda ~4 minutos; después, ~1 segundo.
> Para forzarlo:
> ```python
> from src.apu_source import MatricesProvider
> MatricesProvider().cargar_indice(forzar=True)
> ```

Variables de entorno:

| Variable | Para qué |
|---|---|
| `PONTEM_COSTOS_DB` | ruta exacta de la BD canónica (snapshot o BD viva) |
| `PONTEM_COSTOS_JSON` | ruta exacta del catálogo JSON |
| `PONTEM_COSTOS_SIN_MATRICES` | `1` desactiva el fallback de matrices |

**Cómo se busca y decide cada concepto:**

1. **Descubre el repo entre carpetas hermanas** automáticamente, priorizando el
   que tiene `buscador_costos.py` y el catálogo. Excluye la propia carpeta.
2. **Nivel 3 — buscador oficial:** si existe `buscador_costos.py`, lo importa y
   usa su función `buscar()` como fuente de candidatos. Si no, cae a
   **Nivel 1** (lectura directa del JSON). La app muestra qué nivel usó.
3. Cada concepto de entrada se **re-rankea** con: normalización, stopwords,
   filtro/penalización por unidad y cobertura ponderada por rareza (IDF).
4. Resultado por concepto:
   - **confirmado** — score ≥ umbral (por defecto 55). Precio tomado del catálogo.
   - **dudoso** — score entre el piso (40) y el umbral. Precio del catálogo pero **revisar**.
   - **sin_match** — score < piso. Se muestra en el presupuesto con **PU $0 y fila roja** para completarlo.
5. Las líneas dudosas con un precio anómalo (>12× la mediana de su unidad) se dejan en **$0** para no inflar el total.
6. **Nada se pierde:** el `Presupuesto` incluye todos los conceptos. Las filas rojas son las que debes completar en la hoja `Revisión`.

> Puedes forzar un repo específico pasando su ruta como `repo_path` (app web) o editando la ruta en la CLI.

### Qué PU se cotiza (oficial vs. APU)

El modelo canónico de costos define que **manda el PU publicado** mientras el
APU no esté validado. El presupuesto lo respeta:

| Modo | PU cotizado | Cuándo usarlo |
|---|---|---|
| **Por defecto** | `v_concepto_pu.precio_unitario` (el oficial del catálogo) | Cotizar igual que el catálogo oficial |
| `--cotizar-apu` (CLI) / checkbox en la app | el PU que recalcula la hoja APU | Quiero que editar la hoja `Insumos` mueva el total |

En **ambos** modos la diferencia queda a la vista:

- La hoja **Revisión** trae `PU oficial`, `PU APU`, `Dif. %`, `Origen precio` y
  `Fuente desglose` (resaltada si el desglose no vino de la canónica).
- Cada hoja **APU** muestra el `PU DE VENTA` (recalculado) y el
  `PU OFICIAL (CATALOGO)` uno debajo.
- El **validador** (check 9/9) avisa cuando el PU cotizado difiere del oficial,
  con el impacto en dinero.

> Ejemplo real: `10401-172` tiene PU oficial **$52.97** y su APU calcula
> **$607.72** (+1047%). Con el modo oficial se cotiza $52.97; la diferencia
> queda registrada en la Revisión y en la hoja APU.

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
[1/8] Copiando plantilla...              ✓ machote.xlsx
[2/8] Escribiendo Variables...            ✓ indirectos, utilidad, IVA
[3/8] Preparando APUs...                  ✓ N APUs (BD canónica)
[4/8] Escribiendo Cuadrillas...           ✓ N cuadrillas (IDs locales)
[5/8] Escribiendo Insumos...              ✓ N insumos
[6/8] Creando hojas APU...                ✓ N hojas
[7/8] Escribiendo Presupuesto...          ✓ fórmulas escritas
[8/8] Escribiendo Portada...              ✓ resumen del proyecto
Guardando...                              ✓ ./salidas/.../presupuesto.xlsx
Validando libro...                        ✓ 0 errores, 0 advertencias
[✓] PDF generado                          ✓ Presupuesto + APUs
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
9. El **PU cotizado coincide con el PU oficial del catálogo**: si difiere, se
   reporta como advertencia con el impacto en dinero (se omite si la BD
   canónica no está disponible).

Si el libro tiene errores, además se guarda una copia como
`presupuesto_con_errores.xlsx` en la misma carpeta para depuración.

---

## Plan de implementación

| Fase | Entregable | Criterio |
|---|---|---|
| F0 | Estructura, `config.yaml`, `requirements.txt` | `pip install -r` sin errores |
| F1 | `indexador.py` y `lector_entrada.py` | Carga los 12,927 conceptos y el archivo de entrada |
| F2 | `lector_fasar.py` | Extrae parámetros, categorías y cuadrillas |
| F3 | `matcher.py` sin LLM | ≥90% de la entrada mapeada o marcada |
| F4 | `constructor.py`: hojas base | Abre en Excel/LibreOffice sin errores |
| F5 | `constructor.py`: `Presupuesto` y `Revisión` | Fórmulas `VLOOKUP` funcionan |
| F6 | `constructor.py`: `APU_<id>` | Cambiar un insumo recalcula el PU |
| F7 | `validador.py` | Reporte limpio en caso feliz; detección de errores inyectados |
| F8 | `llm.py` | Sugerencias en `Revisión` sin tocar otras hojas |
| F9 | CLI final y pruebas end-to-end | Corre con archivos de ejemplo reales |