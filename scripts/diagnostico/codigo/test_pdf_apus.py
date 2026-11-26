"""Test completo: build con APUs desglosados + export a PDF."""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[3]))  # raiz del repo
import json, sys
from pathlib import Path

from src.lector_entrada import leer_entrada
from src.indexador import Indexador
from src.matcher import Matcher
from src.constructor import construir_presupuesto
from src.pdf_export import exportar_pdf

# Reutiliza el catálogo mock del test anterior
cat = Path("scripts/diagnostico/datos/catalogo_venado.json")
if not cat.exists():
    print("Falta tests/catalogo_venado.json — corre primero test_venado_e2e.py")
    raise SystemExit(1)

conceptos = leer_entrada(r"H:\Mi unidad\ICEL\PUs\CASA VENADO\VENADO.xlsx")
idx = Indexador(str(cat)); idx.cargar()
m = Matcher(idx, umbral=85)
resultados = m.emparejar(conceptos)

# Marcar los primeros 3 conceptos para desglosar APU
for r in resultados[:3]:
    r["desglosar_apu"] = "S"

proyecto = {
    "nombre": "CASA VENADO", "clave": "VENADO", "cliente": "CLIENTE ICEL",
    "fecha": "2026-11-21", "cotizacion": "ICEL-VENADO-001", "revision": "R02",
    "indirectos_campo": 0.14, "indirectos_oficina": 0.05,
    "financiamiento": 0.0135, "utilidad": 0.08, "iva": 0.16,
}

salida = construir_presupuesto(
    ruta_plantilla="plantillas/machote.xlsx",
    proyecto=proyecto, resultados_matcheo=resultados, indice=idx,
    carpeta_salida="scripts/diagnostico/salidas/vent_venado_apu", usar_timestamp=False,
)
print(f"XLSX: {salida}")

import openpyxl
wb = openpyxl.load_workbook(str(salida))
apus = [s for s in wb.sheetnames if s.upper().startswith("APU")]
print(f"Hojas APU: {apus}")

pdf = exportar_pdf(str(salida))
print(f"PDF: {pdf}")

from pypdf import PdfReader
print(f"Paginas PDF: {len(PdfReader(pdf).pages)}")