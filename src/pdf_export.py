"""
pdf_export.py
=============
Exporta a PDF un libro de Excel mostrando únicamente la hoja `Presupuesto`
y todas las hojas `APU*`, cada una en su propia página.

Estrategia: exporta cada hoja objetivo a un PDF temporal con Excel (COM) y
luego une los PDFs con pypdf. Esto garantiza que cada APU quede en su página.

Requiere: Microsoft Excel (Windows) + pywin32 + pypdf.

Uso:
    from src.pdf_export import exportar_pdf

    pdf = exportar_pdf("salidas/presupuesto.xlsx")
    # -> salidas/presupuesto.pdf
"""

import os
import tempfile
from pathlib import Path


def _hojas_a_imprimir(sheetnames: list[str]) -> list[str]:
    """Selecciona y ordena: Portada, Presupuesto, luego cada APU (por exclusión)."""
    fijas = {"portada", "presupuesto", "parametros", "calculo_categoria",
             "cuadrillas", "variables", "insumos", "notas", "categorias"}
    portada = []
    presupuesto = []
    apus = []
    for nombre in sheetnames:
        n = nombre.strip().lower()
        if n == "portada":
            portada.append(nombre)
        elif n == "presupuesto":
            presupuesto.append(nombre)
        elif n not in fijas:
            apus.append(nombre)
    return portada + presupuesto + apus


def exportar_pdf(ruta_xlsx: str, ruta_pdf: str | None = None,
                 mantener_proceso_abierto: bool = False) -> str:
    """Exporta a PDF solo la hoja Presupuesto + las hojas APU (una por página).

    Args:
        ruta_xlsx: Ruta al archivo .xlsx generado.
        ruta_pdf: Ruta del PDF de salida. Si es None, se usa el mismo nombre .pdf.
        mantener_proceso_abierto: Si True, deja Excel abierto (debug).

    Returns:
        Ruta del PDF generado.
    """
    ruta_xlsx = Path(ruta_xlsx).resolve()
    if not ruta_xlsx.exists():
        raise FileNotFoundError(f"No existe el archivo: {ruta_xlsx}")

    if ruta_pdf is None:
        ruta_pdf = ruta_xlsx.with_suffix(".pdf")
    else:
        ruta_pdf = Path(ruta_pdf).resolve()

    try:
        import win32com.client as win32
    except ImportError as e:
        raise RuntimeError("Se requiere pywin32. Instala: pip install pywin32") from e

    try:
        from pypdf import PdfWriter, PdfReader
    except ImportError as e:
        raise RuntimeError("Se requiere pypdf. Instala: pip install pypdf") from e

    excel = None
    wb = None
    temp_dir = tempfile.mkdtemp(prefix="pontem_pdf_")
    temp_pdfs: list[str] = []

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        wb = excel.Workbooks.Open(str(ruta_xlsx))
        nombres = [s.Name for s in wb.Sheets]
        objetivo = _hojas_a_imprimir(nombres)

        if not objetivo:
            raise RuntimeError("No se encontró 'Presupuesto' ni hojas 'APU'.")

        # Configurar página de cada hoja objetivo
        for nombre in objetivo:
            sh = wb.Sheets(nombre)
            ps = sh.PageSetup
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = False
            ps.Orientation = 2 if nombre.strip().lower() == "presupuesto" else 1
            ps.LeftMargin = excel.InchesToPoints(0.3)
            ps.RightMargin = excel.InchesToPoints(0.3)
            ps.TopMargin = excel.InchesToPoints(0.4)
            ps.BottomMargin = excel.InchesToPoints(0.4)

        try:
            excel.CalculateUntilAsyncQueriesDone()
            excel.CalculateFull()
        except Exception:
            pass

        # Exportar cada hoja a un PDF temporal
        for i, nombre in enumerate(objetivo):
            sh = wb.Sheets(nombre)
            sh.Activate()
            temp_pdf = os.path.join(temp_dir, f"hoja_{i:03d}.pdf")
            sh.ExportAsFixedFormat(
                Type=0, Filename=temp_pdf, Quality=0,
                IncludeDocProperties=True, IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
            if os.path.exists(temp_pdf):
                temp_pdfs.append(temp_pdf)

        if not temp_pdfs:
            raise RuntimeError("Excel no generó ningún PDF temporal.")

        # Unir todos los PDFs
        writer = PdfWriter()
        for tp in temp_pdfs:
            reader = PdfReader(tp)
            for page in reader.pages:
                writer.add_page(page)
        with open(ruta_pdf, "wb") as f:
            writer.write(f)

    except Exception as e:
        raise RuntimeError(f"Error exportando a PDF: {e}") from e

    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None and not mantener_proceso_abierto:
                excel.Quit()
        except Exception:
            pass
        # Limpiar temporales
        for tp in temp_pdfs:
            try:
                os.remove(tp)
            except Exception:
                pass
        try:
            os.rmdir(temp_dir)
        except Exception:
            pass

    return str(ruta_pdf)


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Uso: python src/pdf_export.py <ruta.xlsx> [ruta.pdf]")
        raise SystemExit(1)
    salida = exportar_pdf(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
    print(f"PDF generado: {salida}")