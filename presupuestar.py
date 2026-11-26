"""
presupuestar.py
===============
App única para generar presupuestos de construcción.

Uso:
    python presupuestar.py

Flujo:
    1. Escoger archivo de entrada (.xlsx con el catálogo de conceptos)
    2. Ingresar datos del proyecto y variables de precio
    3. La app indexa, matchea y construye el presupuesto automáticamente
    4. El resultado se guarda en la misma carpeta del archivo origen
"""

import sys
import os
from pathlib import Path
from datetime import date

# Asegurar que src/ está en el path
SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

import typer
from typing import List, Optional
from rich.console import Console
from rich.table import Table
from rich.panel import Panel

from src.lector_entrada import leer_entrada
from src.matcher import Matcher
from src.constructor import construir_presupuesto
from src.costeo_client import ClienteCosteo
from src.apu_db import ProveedorAPU

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

app = typer.Typer(
    name="presupuestar",
    help="Genera un presupuesto de construccion en Excel a partir de un catalogo de conceptos",
    add_completion=False,
)

console = Console()

# Rutas por defecto relativas al script
PLANTILLA_DEFAULT = SCRIPT_DIR / "plantillas" / "machote.xlsx"


def detectar_repo_pontem() -> Path:
    """Busca el repo PONTEM (carpeta con catalogo de costos) entre carpetas hermanas."""
    base = SCRIPT_DIR.parent
    candidatos = [base / "costeo", base / "pontem-costos"]
    try:
        if base.exists():
            for d in base.iterdir():
                if d.is_dir() and d not in candidatos:
                    candidatos.append(d)
    except Exception:
        pass
    for c in candidatos:
        try:
            if (c / "costos_unitarios" / "catalogo_costos.json").exists():
                return c
            if (c / "catalogo_costos.json").exists():
                return c
        except Exception:
            continue
    return base / "costeo"


REPO_PONTEM_DEFAULT = detectar_repo_pontem()

# ---------------------------------------------------------------------------
# Helpers de UI
# ---------------------------------------------------------------------------

def titulo(texto: str):
    console.print()
    console.print(Panel(texto, style="bold white on blue", width=72))


def paso(num: int, texto: str):
    console.print(f"  [[bold cyan]{num}[/]] {texto}")


def ok(texto: str = ""):
    console.print(f"  [bold green][OK][/] {texto}")


def error(texto: str):
    console.print(f"  [bold red][ERROR][/] {texto}")


# ---------------------------------------------------------------------------
# Búsqueda del catálogo PONTEM
# ---------------------------------------------------------------------------

def buscar_catalogo(ruta_base: Path) -> Path | None:
    """Busca catalogo_costos.json en orden de prioridad.

    1. ruta_base/costos_unitarios/catalogo_costos.json (estructura estándar PONTEM)
    2. ruta_base/catalogo_costos.json
    3. Búsqueda recursiva de catalogo_costos.json
    4. Cualquier catalogo*.json (solo si no hay catalogo_costos.json exacto)
    """
    # Prioridad 1: ruta exacta estándar
    candidatos = [
        ruta_base / "costos_unitarios" / "catalogo_costos.json",
        ruta_base / "catalogo_costos.json",
    ]
    for p in candidatos:
        if p.exists():
            return p

    # Prioridad 2: buscar catalogo_costos.json recursivamente
    for p in ruta_base.glob("**/catalogo_costos.json"):
        return p

    # Prioridad 3: solo si no hay catalogo_costos.json, buscar alternativas
    for p in ruta_base.glob("**/catalogo*.json"):
        return p

    return None


# ---------------------------------------------------------------------------
# Comando principal
# ---------------------------------------------------------------------------

@app.command()
def main(
    entrada: Optional[List[str]] = typer.Argument(
        None,
        help="Archivo .xlsx de entrada (si la ruta tiene espacios, no hace falta comillas)",
    ),
    cotizar_apu: bool = typer.Option(
        False,
        "--cotizar-apu",
        help=("Cotiza el PU que recalcula la hoja APU en vez del PU oficial del "
              "catalogo (regla del modelo canonico: publicado > APU > manual). "
              "Con esta opcion, editar Insumos mueve el total."),
    ),
):
    """Genera un presupuesto de construccion en Excel.

    Solo necesitas el archivo de entrada con los conceptos.
    La app busca automaticamente el catalogo PONTEM y la plantilla base.
    El resultado se guarda en la misma carpeta del archivo de entrada.
    """

    titulo("PRESUPUESTADOR PONTEM")
    console.print("  Genera presupuestos de obra en Excel con formulas nativas.\n")
    console.print("  [dim]Columnas esperadas en el archivo de entrada:[/]")
    console.print("  [dim]  id | partida | concepto | unidad | cantidad[/]\n")

    # =====================================================================
    # 1. SELECCIONAR ARCHIVO DE ENTRADA
    # =====================================================================
    if not entrada:
        entrada_str = typer.prompt("  Archivo de entrada (.xlsx)")
    else:
        # Reconstruir la ruta aunque PowerShell/cmd la haya partido por espacios
        entrada_str = " ".join(entrada)
        console.print(f"  [dim]Ruta: {entrada_str}[/]")

    ruta_entrada = Path(entrada_str).resolve()

    # Si es una carpeta, buscar archivos .xlsx dentro
    if ruta_entrada.is_dir():
        xlsx_files = sorted(ruta_entrada.glob("*.xlsx"))
        if not xlsx_files:
            error(f"No se encontraron archivos .xlsx en: {ruta_entrada}")
            raise typer.Exit(code=1)
        if len(xlsx_files) == 1:
            ruta_entrada = xlsx_files[0]
            console.print(f"  [dim]Archivo detectado: {ruta_entrada.name}[/]")
        else:
            console.print(f"  [dim]Se encontraron {len(xlsx_files)} archivos .xlsx:[/]")
            for i, f in enumerate(xlsx_files, 1):
                console.print(f"    [dim]{i}.[/] {f.name}")
            seleccion = typer.prompt(f"  Selecciona (1-{len(xlsx_files)})", type=int)
            if 1 <= seleccion <= len(xlsx_files):
                ruta_entrada = xlsx_files[seleccion - 1]
            else:
                error("Seleccion invalida")
                raise typer.Exit(code=1)

    if not ruta_entrada.exists():
        error(f"No se encontro el archivo: {ruta_entrada}")
        raise typer.Exit(code=1)

    if not ruta_entrada.is_file():
        error(f"La ruta no es un archivo: {ruta_entrada}")
        raise typer.Exit(code=1)

    # Validar que se puede leer
    try:
        conceptos_entrada = leer_entrada(str(ruta_entrada))
    except Exception as e:
        error(f"Error al leer el archivo: {e}")
        raise typer.Exit(code=1)

    carpeta_origen = ruta_entrada.parent
    ok(f"Archivo: {ruta_entrada.name} ({len(conceptos_entrada)} conceptos)")
    console.print(f"  [dim]Salida en: {carpeta_origen}[/]")

    # =====================================================================
    # 2. UBICAR REPO DE COSTOS (contrato de consumo)
    # =====================================================================
    titulo("REPO DE COSTOS PONTEM")

    cost_client = ClienteCosteo()  # descubre entre carpetas hermanas

    if not cost_client.disponible:
        console.print("  [dim]No se encontro el repo de costos en carpetas hermanas.[/]")
        ruta_repo = typer.prompt("  Ruta de la carpeta que contiene catalogo_costos.json")
        cost_client = ClienteCosteo(ruta_repo)

    if not cost_client.disponible or cost_client.indice_path is None:
        error("No se encontro el catalogo de costos. Verifica la ruta del repo.")
        raise typer.Exit(code=1)

    console.print(f"  [dim]Repo: {cost_client.repo}[/]")
    ok(f"Contrato: {cost_client.descripcion_contrato()}")

    # Proveedor de APUs (BD canónica de costos; fallback de matrices si falta)
    apu_prov = ProveedorAPU(repo=str(cost_client.repo) if cost_client.repo else None)
    if apu_prov.disponible:
        n_apus = apu_prov.cargar_indice()
        stats_prov = apu_prov.stats().get("canonica", {})
        ok(f"BD canónica de costos: {n_apus:,} APUs "
           f"({stats_prov.get('conceptos', 0):,} conceptos)")
    else:
        console.print("  [bold yellow][!][/] No se encontro la BD canónica de costos "
                      "ni las matrices; las hojas APU e Insumos quedarán vacías.")

    # =====================================================================
    # 3. DATOS DEL PROYECTO Y VARIABLES
    # =====================================================================
    titulo("DATOS DEL PROYECTO")

    nombre_proy = typer.prompt("  Nombre del proyecto")
    cliente = typer.prompt("  Cliente / Empresa")
    fecha_proy = typer.prompt("  Fecha (YYYY-MM-DD)", default=date.today().isoformat())
    cotizacion = typer.prompt("  Cotización No.", default="COT-001")
    revision = typer.prompt("  Revisión", default="R01")

    console.print()
    console.print("  [bold]Variables de precio y matcheo:[/]")
    ind_campo = float(typer.prompt("    % Indirectos de campo", default="0.14"))
    ind_oficina = float(typer.prompt("    % Indirectos de oficina", default="0.05"))
    financiamiento = float(typer.prompt("    % Financiamiento", default="0.0135"))
    utilidad = float(typer.prompt("    % Utilidad", default="0.08"))
    iva = float(typer.prompt("    % IVA", default="0.16"))
    console.print("    [dim]Umbral: score minimo para marcar un match como 'confirmado'.[/]")
    console.print("    [dim]  55-70: recomendado (buen balance)[/]")
    console.print("    [dim]  >80: estricto (marca mas como dudoso, pero el presupuesto igual se llena)[/]")
    umbral_fuzzy = int(typer.prompt("    Umbral (0-100)", default="55"))

    proyecto = {
        "nombre": nombre_proy,
        "clave": ruta_entrada.stem,
        "cliente": cliente,
        "fecha": fecha_proy,
        "cotizacion": cotizacion,
        "revision": revision,
        "indirectos_campo": ind_campo,
        "indirectos_oficina": ind_oficina,
        "financiamiento": financiamiento,
        "utilidad": utilidad,
        "iva": iva,
    }

    # =====================================================================
    # 4. MATCHEO AUTOMÁTICO
    # =====================================================================
    titulo("MATCHEO DE CONCEPTOS")
    console.print("  [dim]Comparando contra el catalogo PONTEM...[/]\n")

    paso(1, "Matcheando conceptos...")
    matcher = Matcher(cliente=cost_client, umbral=umbral_fuzzy, piso=40)
    resultados = matcher.emparejar(conceptos_entrada)
    stats = matcher.stats_matcheo(resultados)
    ok(f"{stats['exactos']} exactos, {stats['oficiales']} oficiales, "
       f"{stats['fuzzy']} fuzzy, {stats['llm']} IA, {stats['sin_match']} sin match "
       f"({stats['pct_confirmados']}% confirmados)")

    # ---- IA opcional para dudosos / sin match ----
    try:
        from src.llm import cargar_config_llm, ClienteLLM, aplicar_sugerencias
        llm_cfg = cargar_config_llm(str(SCRIPT_DIR / "config.yaml"))
        llm = ClienteLLM(llm_cfg)
        pendientes_ia = sum(1 for r in resultados if r["estado"] in ("dudoso", "sin_match"))
        if pendientes_ia:
            console.print(f"  [dim]{llm.estado()}[/]")
            usar_ia = typer.confirm(
                f"  ¿Usar IA para resolver {pendientes_ia} conceptos dudosos/sin match?",
                default=False,
            )
            if usar_ia:
                if not llm.disponible:
                    console.print("  [bold yellow][!][/] IA no configurada. Copia .env.example a .env "
                                  "y pon tu DEEPSEEK_API_KEY (o OPENROUTER_API_KEY).")
                else:
                    paso(2, f"Consultando IA ({pendientes_ia} conceptos)...")
                    sugerencias = llm.resolver(resultados)
                    # `indice` debe exponer por_clave(): el ClienteCosteo cumple
                    # ese contrato (igual que el flujo web).
                    aplicadas = aplicar_sugerencias(resultados, cost_client, sugerencias,
                                                    auto_aplicar=llm_cfg.get("auto_aplicar", True))
                    ok(f"IA revisó {len(sugerencias)} conceptos, aplicó {aplicadas}")
                    stats = matcher.stats_matcheo(resultados)
    except Exception as e:
        console.print(f"  [bold yellow][!][/] IA no disponible: {e}")

    # Mostrar tabla de revisión
    if resultados:
        console.print()
        table = Table(title="Resultados del matcheo", show_lines=False, width=100)
        table.add_column("ID", style="dim", width=10, no_wrap=True)
        table.add_column("Concepto", width=38)
        table.add_column("Und", width=5)
        table.add_column("Cant", width=10, justify="right")
        table.add_column("Clave PONTEM", style="cyan", width=14)
        table.add_column("Score", width=6, justify="right")
        table.add_column("Estado", width=12)

        for r in resultados:
            estado_color = {
                "confirmado": "[bold green]confirmado[/]",
                "confirmado_ia": "[bold green]confirmado_ia[/]",
                "dudoso": "[bold yellow]dudoso[/]",
                "sin_match": "[bold red]sin_match[/]",
            }
            table.add_row(
                r["id_entrada"],
                r["concepto_entrada"][:36],
                r["unidad_entrada"],
                f"{r['cantidad']:,.1f}",
                r["clave_sugerida"] or "-",
                f"{r['score']:.0f}",
                estado_color.get(r["estado"], r["estado"]),
            )

        console.print(table)

    if stats["dudosos"] > 0 or stats["sin_match"] > 0:
        console.print("\n  [bold yellow][!][/] Los conceptos dudosos quedan marcados en la hoja [bold]Revision[/]")
        console.print("  para que los ajustes manualmente en Excel.")

    # =====================================================================
    # 5. CONFIRMAR Y CONSTRUIR
    # =====================================================================
    console.print()
    confirmar = typer.confirm("  Construir presupuesto?", default=True)

    if not confirmar:
        console.print("\n  [bold yellow]Cancelado.[/]")
        raise typer.Exit()

    console.print()

    # Verificar plantilla
    if not PLANTILLA_DEFAULT.exists():
        error(f"No se encontro la plantilla: {PLANTILLA_DEFAULT}")
        raise typer.Exit(code=1)

    # Construir
    paso(2, "Construyendo presupuesto...")
    ruta_salida = construir_presupuesto(
        ruta_plantilla=str(PLANTILLA_DEFAULT),
        proyecto=proyecto,
        resultados_matcheo=resultados,
        carpeta_salida=str(carpeta_origen),  # misma carpeta que el archivo origen
        usar_timestamp=False,
        apu_provider=apu_prov,
        incluir_revision=False,
        validar=True,
        recalcular=True,
        usar_pu_oficial=not cotizar_apu,
    )
    ok("Presupuesto construido")
    if cotizar_apu:
        console.print("  [dim]PU cotizado: el que recalcula la hoja APU[/]")
    else:
        console.print("  [dim]PU cotizado: el oficial del catalogo (modelo canonico)[/]")

    # =====================================================================
    # 6. EXPORTAR PDF (Presupuesto + cada APU)
    # =====================================================================
    ruta_pdf = None
    try:
        from src.pdf_export import exportar_pdf
        paso(3, "Generando PDF (Presupuesto + APUs)...")
        ruta_pdf = exportar_pdf(str(ruta_salida))
        ok(f"PDF generado: {ruta_pdf}")
    except Exception as e:
        console.print(f"  [bold yellow][!][/] No se pudo generar el PDF: {e}")

    # =====================================================================
    # 7. RESULTADO
    # =====================================================================
    titulo("PRESUPUESTO GENERADO")

    import openpyxl
    wb = openpyxl.load_workbook(str(ruta_salida))
    fijas = {"Portada", "Presupuesto", "Parametros", "Calculo_categoria",
             "Cuadrillas", "Variables", "Insumos", "Notas", "Categorias"}
    hojas_dinamicas = [s for s in wb.sheetnames if s not in fijas]
    apus = hojas_dinamicas

    console.print(f"  [bold]XLSX:[/] {ruta_salida}")
    console.print(f"  [bold]Tamaño:[/] {os.path.getsize(ruta_salida)/1024:.1f} KB")
    console.print(f"  [bold]Hojas:[/] {len(wb.sheetnames)} ({', '.join(hojas_dinamicas)} + FASAR)")
    if ruta_pdf:
        console.print(f"  [bold]PDF:[/]  {ruta_pdf}  ({len(apus)} APU(s) incluidos)")
    console.print()
    console.print("  [bold green]Listo.[/] Abre el XLSX o el PDF. Al modificar un PU en")
    console.print("  la hoja [bold]Insumos[/], todo el presupuesto se recalcula solo.")
    console.print()


# ---------------------------------------------------------------------------
# Entrada
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app()