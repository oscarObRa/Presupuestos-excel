# Scripts de diagnóstico

Scripts auxiliares para **inspeccionar y depurar** el sistema: prints,
comparaciones, previews. **No son la suite de pruebas** (esa vive en
`tests/`) y **no se ejecutan en CI**.

> ⚠️ Estos scripts se movieron aquí el 2026-11-26 y **pueden estar
> desactualizados** respecto al código actual (usan rutas absolutas a
> `C:\Pontem_dev\costeo`, a `G:\Mi unidad\Costos\PU2026\` o a
> `H:\Mi unidad\ICEL\...`, y algunos dependían de la API anterior).
> Si necesitas uno, úsalo como punto de partida y ajústalo; no es un contrato.

## Estructura

| Carpeta | Contenido |
|---|---|
| `codigo/` | los scripts (`.py`) |
| `datos/` | catálogos y entradas de ejemplo (`.json`, `.xlsx`) |
| `salidas/` | salidas generadas (PDF, PNG, `.xlsx`): **ignorada por git** |

## Cómo se ejecutan

Desde la **raíz del repo** (las rutas relativas de datos y salidas lo asumen):

```powershell
venv\Scripts\python.exe scripts\diagnostico\codigo\<script>.py
```

Los scripts que importan `src.*` llevan un bootstrap que agrega la raíz del
repo al `sys.path`:

```python
sys.path.insert(0, str(Path(__file__).resolve().parents[3]))
```

## Qué hay aquí

### Catálogo y matcheo

| Script | Qué hace |
|---|---|
| `diag_top.py` | Top de conceptos por precio unitario del catálogo |
| `top_importes.py` | Top de importes de un presupuesto |
| `debug_scoring.py` | Desglose del score de un matcheo |
| `diagnostico_match.py` | Diagnóstico del buscador oficial (`buscador_costos.py`) |
| `experimento_match.py` | Experimentos de ranking/score |
| `analizar_venado.py` | Análisis del presupuesto VENADO |
| `inspec_claves.py`, `inspec_bloque.py`, `inspec_caratula.py`, `inspec_fuente.py`, `inspec_matriz.py` | Inspección de claves, bloques, carátulas, fuentes y matrices |

### Lector de entrada y partidas

| Script | Qué hace |
|---|---|
| `test_paso3.py` | Genera una entrada de ejemplo, la lee y la valida |
| `test_partidas.py` | Lectura del archivo real VENADO con partidas |
| `test_subpartidas.py` | Subpartidas: genera, lee y construye |
| `revisar_partida03.py` | Revisa la partida 03 |
| `verificar_partidas.py`, `verificar_formato.py`, `verificar_nuevo.py`, `verificar_salida.py` | Verificaciones de formato y de salida |

### Construcción y PDF

| Script | Qué hace |
|---|---|
| `test_pdf_apus.py` | Exporta el PDF con APUs |
| `test_venado_e2e.py` | Flujo completo con VENADO |
| `test_revision_ia.py` | Revisión con sugerencias de IA |
| `test_llm.py` | Prueba del cliente LLM |
| `test_parser_matriz.py` | Parseo de una matriz de `PU2026` |
| `check_db.py` | Inspección de `costos.db` (BD legacy; ya no se usa en el flujo) |
| `ver_g.py`, `inspeccionar_mockup*.py` | Previews y mockups |

## Si agregas un script nuevo

- Va en `codigo/`, no en `tests/`.
- Si importa `src.*`, incluye el bootstrap.
- Sus datos van en `datos/` y sus salidas en `salidas/`.
