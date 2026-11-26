# Suite de pruebas

Pruebas **autocontenidas**: no necesitan el repo de costeos, ni la BD
canónica, ni los Excel de `G:\Mi unidad\Costos\PU2026\`. Todo se genera en
memoria con `fixtures.py`, así que corren igual en CI, en cualquier equipo y
sin acceso a las unidades de Drive.

## Ejecutar

```powershell
venv\Scripts\python.exe tests\run_all.py
```

Es lo que ejecuta el CI. Sale con código `!= 0` si alguna suite falla.

## Suites

| Archivo | Qué cubre |
|---|---|
| `fixtures.py` | datos de prueba en memoria (catálogo, APU, entrada, proyecto) + `construir_de_prueba()` |
| `test_matcher.py` | emparejamiento de conceptos: exacto, mismo nombre en unidades distintas, fuzzy |
| `test_constructor.py` | construcción del libro: Portada, Presupuesto, hojas APU, IDs locales, cuadrillas, **PU oficial vs PU del APU**, trazabilidad del precio en la Revisión |
| `test_validador.py` | el validador **detecta** los defectos: hoja inexistente, clave inexistente, ID local duplicado, unidad inconsistente, totales rotos, fórmula malformada, ciclo entre hojas, cuadrilla en $0, **PU divergente del oficial** |

## Fixtures

`fixtures.py` reproduce el contrato del proveedor canónico: `apu_de()`
devuelve el desglose y la trazabilidad del precio (`pu_oficial`, `pu_apu`,
`origen_precio`, `diferencia_publicado_apu`, `fuente`), y `concepto()` el PU
oficial. Así los tests pueden verificar el PU oficial sin tocar la BD real.

El catálogo de prueba incluye a propósito un caso de **mismo nombre en
unidades distintas** (`CAT-003` en `M2` y `CAT-004` en `TON`), que fue el que
produjo precios 1000× mayores.

## Añadir una prueba

1. Escribe la función en el archivo de suite correspondiente, con `check(...)`
   (los failures se acumulan y se reportan al final).
2. Agrégala a la lista `main()` del archivo.
3. Si necesitas un caso nuevo, agrégalo a `fixtures.py` (sin dependencias
   externas).

## Lo que **no** está aquí

Los scripts de inspección y depuración están en
[`scripts/diagnostico/`](../scripts/diagnostico/README.md). Se movieron aquí el
2026-11-26 para que `tests/` contenga solo la suite que corre en CI.
