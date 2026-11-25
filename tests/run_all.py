"""
run_all.py
==========
Ejecuta toda la suite de pruebas y devuelve un código de salida != 0 si
alguna falla. Es lo que ejecuta el CI.

Ejecutar:  python tests\\run_all.py
"""

import subprocess
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
SUITES = [
    "test_matcher.py",
    "test_constructor.py",
    "test_validador.py",
]


def main():
    print("=" * 60)
    print("SUITE DE PRUEBAS — Presupuestos-excel")
    print("=" * 60)

    fallos = []
    for suite in SUITES:
        ruta = Path(__file__).resolve().parent / suite
        print(f"\n>>> {suite}")
        proc = subprocess.run([sys.executable, str(ruta)], cwd=str(RAIZ))
        if proc.returncode != 0:
            fallos.append(suite)

    print("\n" + "=" * 60)
    if fallos:
        print(f"FALLARON {len(fallos)} suites: {', '.join(fallos)}")
        return 1
    print("TODAS LAS SUITES PASARON")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
