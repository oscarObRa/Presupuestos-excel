"""
app.py — Frontend web con Streamlit para el Presupuestador PONTEM.

Uso:
    streamlit run app.py
"""

import sys
import os
import tempfile
import shutil
from pathlib import Path
from datetime import date

# Asegurar que src/ está en el path
SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))


def detectar_repo_pontem() -> Path:
    """Busca el repo de costos entre carpetas hermanas usando el contrato."""
    try:
        from src.costeo_client import descubrir_repo_costeo
        repo = descubrir_repo_costeo(SCRIPT_DIR)
        if repo:
            return repo
    except Exception:
        pass
    return SCRIPT_DIR.parent / "costeo"


REPO_PONTEM_DEFAULT = detectar_repo_pontem()

import streamlit as st
import pandas as pd
import openpyxl

from src.lector_entrada import leer_entrada
from src.matcher import Matcher
from src.constructor import construir_presupuesto
from src.apu_db import ProveedorAPU
from src.costeo_client import ClienteCosteo

# ---------------------------------------------------------------------------
# Config de página
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Presupuestador PONTEM",
    page_icon="📊",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Sidebar: datos del proyecto
# ---------------------------------------------------------------------------

st.title("📊 Presupuestador PONTEM")
st.caption("Genera presupuestos de construcción en Excel con fórmulas nativas.")

with st.sidebar:
    st.header("⚙️ Datos del Proyecto")

    nombre = st.text_input("Nombre del proyecto", value="PROYECTO DE CONSTRUCCIÓN")
    cliente = st.text_input("Cliente / Empresa", value="CLIENTE, S.A. DE C.V.")
    fecha_proy = st.date_input("Fecha del presupuesto", value=date.today())

    colc1, colc2 = st.columns(2)
    with colc1:
        cotizacion = st.text_input("Cotización No.", value="COT-001")
    with colc2:
        revision = st.text_input("Revisión", value="R01")

    st.divider()
    st.subheader("Variables de precio")

    col1, col2 = st.columns(2)
    with col1:
        ind_campo = st.number_input("% Indirectos campo", min_value=0.0, max_value=1.0,
                                     value=0.14, step=0.01, format="%.2f")
        ind_oficina = st.number_input("% Indirectos oficina", min_value=0.0, max_value=1.0,
                                       value=0.05, step=0.01, format="%.2f")
        financiamiento = st.number_input("% Financiamiento", min_value=0.0, max_value=1.0,
                                          value=0.0135, step=0.001, format="%.4f")
    with col2:
        utilidad = st.number_input("% Utilidad", min_value=0.0, max_value=1.0,
                                    value=0.08, step=0.01, format="%.2f")
        iva = st.number_input("% IVA", min_value=0.0, max_value=1.0,
                              value=0.16, step=0.01, format="%.2f")
        umbral = st.slider("Umbral de match", min_value=30, max_value=90,
                           value=55, help="Score mínimo para marcar un match como 'confirmado'. 55-70 recomendado.")

    st.divider()
    st.subheader("🤖 IA (opcional)")
    try:
        from src.llm import cargar_config_llm, ClienteLLM
        _llm_cfg = cargar_config_llm(str(SCRIPT_DIR / "config.yaml"))
        _llm = ClienteLLM(_llm_cfg)
        st.caption(_llm.estado())
        usar_ia = st.checkbox("Usar IA para dudosos / sin match", value=False,
                              disabled=not _llm.disponible,
                              help="Resuelve los conceptos dudosos con DeepSeek/OpenRouter. Requiere .env con API key.")
    except Exception:
        _llm_cfg, _llm, usar_ia = {}, None, False

    st.divider()
    st.subheader("📁 Catálogo PONTEM")

    repo_path = st.text_input(
        "Ruta del repo PONTEM",
        value=str(REPO_PONTEM_DEFAULT),
        help="Carpeta que contiene costos_unitarios/catalogo_costos.json"
    )

    cotizar_apu = st.checkbox(
        "PU con fórmulas vivas (editar Insumos mueve el total)",
        value=True,
        help=("El PU nace del catálogo de costeo (semilla), pero en el libro "
              "queda ligado con fórmula a la hoja APU (que a su vez lee "
              "Insumos): si luego editas precios en Excel, el PU y el total "
              "se recalculan solos. Desactívalo solo si necesitas congelar "
              "el PU oficial del catálogo como valor fijo."),
    )

    st.divider()
    st.subheader("📂 Carpeta de salida")
    st.caption("El presupuesto se guardará en esta carpeta.")
    
    # Carpeta de salida: por defecto salidas/ junto al script, pero editable
    default_salida = SCRIPT_DIR / "salidas"
    salida_path = st.text_input(
        "Ruta de salida",
        value=str(default_salida),
        help="Carpeta donde se guardará el presupuesto.xlsx y presupuesto.pdf"
    )
    carpeta_salida = Path(salida_path)
    carpeta_salida.mkdir(parents=True, exist_ok=True)
    st.caption(f"📁 Se guardará en: `{carpeta_salida}`")

# ---------------------------------------------------------------------------
# Área principal: archivo de entrada
# ---------------------------------------------------------------------------

st.header("📄 Archivo de entrada")

st.markdown(
    "Sube tu archivo `.xlsx` con los conceptos a cotizar. "
    "Columnas esperadas: `id` | `partida` | `concepto` | `unidad` | `cantidad`"
)

archivo = st.file_uploader(
    "Selecciona el archivo de entrada",
    type=["xlsx"],
    help="También puedes arrastrar el archivo aquí",
)

# Previsualizar archivo si ya se subió
if archivo is not None:
    try:
        df_preview = pd.read_excel(archivo)
        st.write(f"**{len(df_preview)} conceptos encontrados**")
        st.dataframe(df_preview, use_container_width=True, height=200)
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")

# ---------------------------------------------------------------------------
# Estado de sesión para regenerar
# ---------------------------------------------------------------------------
if "presupuesto_generado" not in st.session_state:
    st.session_state.presupuesto_generado = False
    st.session_state.ruta_salida = None
    st.session_state.ruta_pdf = None
    st.session_state.archivo_nombre = None

# ---------------------------------------------------------------------------
# Botón: Generar / Regenerar
# ---------------------------------------------------------------------------

st.divider()

# Determinar qué botón mostrar
if st.session_state.presupuesto_generado and archivo is not None and archivo.name == st.session_state.archivo_nombre:
    col_btn1, col_btn2, col_space = st.columns([1, 1, 2])
    with col_btn1:
        regenerar = st.button("🔄 Regenerar Presupuesto", type="secondary", use_container_width=True,
                              help="Borra el anterior y genera uno nuevo")
    with col_btn2:
        if st.button("🗑️ Limpiar", use_container_width=True):
            st.session_state.presupuesto_generado = False
            st.session_state.ruta_salida = None
            st.session_state.ruta_pdf = None
            st.session_state.archivo_nombre = None
            st.rerun()
else:
    col_btn, col_space = st.columns([1, 3])
    with col_btn:
        generar = st.button("🚀 Generar Presupuesto", type="primary", use_container_width=True)
    regenerar = False

if generar or regenerar:
    if archivo is None:
        st.error("Primero selecciona un archivo de entrada.")
    else:
        # Guardar archivo subido a temporal
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(archivo.getvalue())
            ruta_entrada = tmp.name

        # Usar carpeta de salida del sidebar
        carpeta_salida = Path(salida_path)
        carpeta_salida.mkdir(parents=True, exist_ok=True)

        # Si es regenerar, borrar archivos previos
        if regenerar and st.session_state.ruta_salida:
            try:
                if os.path.exists(st.session_state.ruta_salida):
                    os.unlink(st.session_state.ruta_salida)
                if st.session_state.ruta_pdf and os.path.exists(st.session_state.ruta_pdf):
                    os.unlink(st.session_state.ruta_pdf)
            except Exception:
                pass

        try:
            with st.status("Procesando...", expanded=True) as status:
                # 1. Leer entrada
                st.write("Leyendo archivo de entrada...")
                conceptos_entrada = leer_entrada(ruta_entrada)
                st.write(f"✓ {len(conceptos_entrada)} conceptos leídos")

                # 2. Ubicar repo de costos (contrato de consumo)
                st.write("Ubicando repo de costos PONTEM...")
                from src.costeo_client import ClienteCosteo
                cost_client = ClienteCosteo(repo_path) if repo_path else ClienteCosteo()

                if not cost_client.disponible or cost_client.indice_path is None:
                    st.error(f"No se encontró el catálogo de costos en `{repo_path}`")
                    st.info("Asegúrate de que la ruta apunte al repo de costos (buscador_costos.py / costos_unitarios/catalogo_costos.json).")
                    st.stop()

                st.write(f"✓ {cost_client.descripcion_contrato()}")

                # Proveedor de APUs (BD canónica de costos; fallback matrices)
                from src.apu_db import ProveedorAPU
                apu_prov = ProveedorAPU(repo=str(cost_client.repo) if cost_client.repo else None)
                if apu_prov.disponible:
                    n_apus = apu_prov.cargar_indice()
                    stats_prov = apu_prov.stats().get("canonica", {})
                    st.write(f"✓ BD canónica de costos: {n_apus:,} APUs "
                             f"({stats_prov.get('conceptos', 0):,} conceptos)")
                else:
                    st.warning("No se encontró la BD canónica de costos; APUs/Insumos quedarán vacíos.")

                # 3. Matchear (usa ClienteCosteo.buscar() Nivel 3 como fuente principal)
                st.write("Matcheando conceptos...")
                matcher = Matcher(cliente=cost_client, umbral=umbral)
                resultados = matcher.emparejar(conceptos_entrada)
                stats = matcher.stats_matcheo(resultados)

                # 3b. IA opcional para dudosos / sin match
                if usar_ia and _llm is not None:
                    pend = sum(1 for r in resultados if r["estado"] in ("dudoso", "sin_match"))
                    if pend:
                        st.write(f"Consultando IA ({pend} conceptos dudosos/sin match)...")
                        try:
                            from src.llm import aplicar_sugerencias
                            sugerencias = _llm.resolver(resultados)
                            aplicadas = aplicar_sugerencias(resultados, cost_client, sugerencias,
                                                            auto_aplicar=_llm_cfg.get("auto_aplicar", True))
                            st.write(f"✓ IA revisó {len(sugerencias)}, aplicó {aplicadas}")
                            stats = matcher.stats_matcheo(resultados)
                        except Exception as eia:
                            st.warning(f"IA: {eia}")

                # Mostrar tabla de resultados
                st.write(f"✓ {stats['exactos']} exactos, {stats['fuzzy']} fuzzy, "
                         f"{stats['sin_match']} sin match ({stats['pct_confirmados']}%)")

                df_resultados = pd.DataFrame(resultados)
                columnas_mostrar = ["id_entrada", "concepto_entrada", "unidad_entrada",
                                    "cantidad", "clave_sugerida", "score", "estado"]
                st.dataframe(
                    df_resultados[columnas_mostrar],
                    use_container_width=True,
                    height=300,
                    column_config={
                        "estado": st.column_config.TextColumn(
                            "Estado",
                            help="confirmado = match exitoso | dudoso = revisar | sin_match = no encontrado"
                        ),
                        "score": st.column_config.ProgressColumn(
                            "Score",
                            min_value=0,
                            max_value=100,
                            format="%.0f",
                        ),
                    },
                )

                # 4. Construir
                st.write("Construyendo presupuesto...")
                proyecto = {
                    "nombre": nombre,
                    "clave": archivo.name.replace(".xlsx", ""),
                    "cliente": cliente,
                    "fecha": str(fecha_proy),
                    "cotizacion": cotizacion,
                    "revision": revision,
                    "indirectos_campo": ind_campo,
                    "indirectos_oficina": ind_oficina,
                    "financiamiento": financiamiento,
                    "utilidad": utilidad,
                    "iva": iva,
                }

                plantilla = SCRIPT_DIR / "plantillas" / "machote.xlsx"

                ruta_salida = construir_presupuesto(
                    ruta_plantilla=str(plantilla),
                    proyecto=proyecto,
                    resultados_matcheo=resultados,
                    carpeta_salida=str(carpeta_salida),
                    usar_timestamp=True,
                    apu_provider=apu_prov,
                    incluir_revision=False,
                    # Validación estática (rápida). El cuadre de totales con
                    # recálculo de Excel se omite en la web: el PDF ya abre Excel.
                    validar=True,
                    recalcular=False,
                    usar_pu_oficial=not cotizar_apu,
                )
                st.write(f"✓ Presupuesto generado")

                # Generar PDF (solo Presupuesto + cada APU)
                st.write("Generando PDF (Presupuesto + APUs)...")
                try:
                    from src.pdf_export import exportar_pdf
                    ruta_pdf = exportar_pdf(str(ruta_salida))
                    st.write("✓ PDF generado")
                except Exception as epdf:
                    ruta_pdf = None
                    st.warning(f"No se pudo generar el PDF: {epdf}")

                status.update(label="Procesamiento completo!", state="complete")

            # Guardar en session state para regeneración y descarga
            st.session_state.presupuesto_generado = True
            st.session_state.ruta_salida = str(ruta_salida)
            st.session_state.ruta_pdf = str(ruta_pdf) if ruta_pdf else None
            st.session_state.archivo_nombre = archivo.name

            # 5. Mostrar resultado y descarga
            st.success(f"✅ Presupuesto generado exitosamente")

            st.info(f"📁 Archivo guardado en: `{ruta_salida}`")
            if ruta_pdf:
                st.info(f"📄 PDF guardado en: `{ruta_pdf}`")

            # Leer archivo para descarga (el download_button descarga al navegador, pero el archivo ya está en la carpeta de salida)
            colx, colp = st.columns(2)
            with colx:
                with open(ruta_salida, "rb") as f:
                    st.download_button(
                        label="📥 Descargar presupuesto.xlsx",
                        data=f,
                        file_name=f"presupuesto_{archivo.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )
            with colp:
                if ruta_pdf and os.path.exists(ruta_pdf):
                    with open(ruta_pdf, "rb") as f:
                        st.download_button(
                            label="🖨️ Descargar PDF (Presupuesto + APUs)",
                            data=f,
                            file_name=f"presupuesto_{archivo.name.replace('.xlsx', '')}.pdf",
                            mime="application/pdf",
                            type="secondary",
                            use_container_width=True,
                        )

            # Info de hojas
            wb = openpyxl.load_workbook(str(ruta_salida))
            fijas = {"Portada", "Presupuesto", "Parametros", "Calculo_categoria",
                     "Cuadrillas", "Variables", "Insumos", "Notas", "Categorias"}
            hojas_dinamicas = [s for s in wb.sheetnames if s not in fijas]
            apus = hojas_dinamicas
            st.caption(f"Hojas generadas: {', '.join(hojas_dinamicas)}")
            st.caption(f"APUs en PDF: {len(apus)} hoja(s) individual(es)")
            st.caption("💡 Abre el archivo en Excel. Al modificar un PU en **Insumos**, todo se recalcula automáticamente.")

            # Limpiar temporal
            os.unlink(ruta_entrada)

        except Exception as e:
            st.error(f"Error: {e}")
            import traceback
            st.code(traceback.format_exc())

# ---------------------------------------------------------------------------
# Mostrar resultado previo si ya se generó (y no es una nueva generación)
# ---------------------------------------------------------------------------
elif st.session_state.presupuesto_generado and archivo is not None and archivo.name == st.session_state.archivo_nombre:
    st.success(f"✅ Presupuesto ya generado: `{st.session_state.archivo_nombre}`")
    
    ruta_salida = st.session_state.ruta_salida
    ruta_pdf = st.session_state.ruta_pdf
    
    st.info(f"📁 Archivo guardado en: `{ruta_salida}`")
    if ruta_pdf:
        st.info(f"📄 PDF guardado en: `{ruta_pdf}`")

    # Botones de descarga
    colx, colp = st.columns(2)
    with colx:
        if os.path.exists(ruta_salida):
            with open(ruta_salida, "rb") as f:
                st.download_button(
                    label="📥 Descargar presupuesto.xlsx",
                    data=f,
                    file_name=f"presupuesto_{archivo.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
        else:
            st.warning("⚠️ Archivo XLSX no encontrado en disco. Haz clic en 'Regenerar'.")
    with colp:
        if ruta_pdf and os.path.exists(ruta_pdf):
            with open(ruta_pdf, "rb") as f:
                st.download_button(
                    label="🖨️ Descargar PDF (Presupuesto + APUs)",
                    data=f,
                    file_name=f"presupuesto_{archivo.name.replace('.xlsx', '')}.pdf",
                    mime="application/pdf",
                    type="secondary",
                    use_container_width=True,
                )
        elif ruta_pdf:
            st.warning("⚠️ Archivo PDF no encontrado en disco.")

    # Info de hojas
    if os.path.exists(ruta_salida):
        wb = openpyxl.load_workbook(str(ruta_salida))
        fijas = {"Portada", "Presupuesto", "Parametros", "Calculo_categoria",
                 "Cuadrillas", "Variables", "Insumos", "Notas", "Categorias"}
        hojas_dinamicas = [s for s in wb.sheetnames if s not in fijas]
        apus = hojas_dinamicas
        st.caption(f"Hojas generadas: {', '.join(hojas_dinamicas)}")
        st.caption(f"APUs en PDF: {len(apus)} hoja(s) individual(es)")
        st.caption("💡 Abre el archivo en Excel. Al modificar un PU en **Insumos**, todo se recalcula automáticamente.")

# ---------------------------------------------------------------------------
# Pie
# ---------------------------------------------------------------------------

st.divider()
st.caption("Presupuestador PONTEM · Fórmulas FASAR (RLOPSRM art. 190-192) · Catálogo de costos unitarios")