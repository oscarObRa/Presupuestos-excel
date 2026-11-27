"""
constructor.py
==============
Construye el archivo .xlsx de salida a partir de la plantilla base,
los datos del proyecto y los resultados del matcheo.

Flujo:
1. Copiar plantillas/machote.xlsx → salidas/<timestamp>/presupuesto.xlsx
2. Llenar Variables con datos del proyecto
3. Escribir Insumos (tabla maestra con fórmulas)
4. Escribir Revision (resultados del matcheo)
5. Crear APU_<id> dinámicos (para conceptos desglosados)
6. Llenar Presupuesto con VLOOKUP → Insumos
7. Crear Explosion (BOM — explosión de insumos)
"""

import os
import re
import shutil
import datetime
from copy import copy
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter as _xl_col_letter
from openpyxl.drawing.image import Image as XLImage

from src.visual_dna import get_dna
from src.formulas import (
    VLOOKUP, VLOOKUP_col, SUMIF, SUMPRODUCT, SUM,
    ref_hoja, celda, rango,
    formula_producto, formula_total_con_indirectos,
)
from src.estilos import (
    escribir_celda, col_letra, aplicar_borde, merge_y_escribir,
    font_titulo, font_titulo_blanco, font_subtitulo, font_header_blanco,
    font_header_negro, font_header_col, font_seccion, font_normal, font_normal_rojo,
    font_negrita, font_italica, font_valor_editable, font_apu_titulo, font_apu_concepto,
    font_subtotal, font_param_header,
    fill_azul_oscuro, fill_azul_header, fill_azul_medio, fill_gris_seccion,
    fill_blanco, fill_verde_claro, fill_amarillo_claro, fill_pu_venta,
    fill_negro, fill_azul_presupuesto,
    align_center, align_left, align_right, align_center_nowrap,
    AZUL_OSCURO, AZUL_HEADER, AZUL_MEDIO, GRIS_SECCION, NEGRO_TEXTO, ROJO_TEXTO,
    VERDE_CLARO, AMARILLO_CLARO, AZUL_PU_VENTA,
)
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage


# =============================================================================
# Constructor principal
# =============================================================================

class Constructor:
    """Construye el presupuesto .xlsx completo."""

    def __init__(self, ruta_plantilla: str, carpeta_salida: str = "salidas",
                 usar_timestamp: bool = False, apu_provider=None,
                 incluir_revision: bool = False,
                 usar_pu_oficial: bool = False,
                 empresa: Optional[str] = None):
        # ADN visual por empresa: la identidad (colores, retícula, assets,
        # membrete) se resuelve aquí y NUNCA se hardcodea en los métodos.
        # Con empresa=None se usa la empresa por defecto del ADN (ICEL).
        self.dna = get_dna(empresa)
        self.empresa = empresa
        self.ruta_plantilla = Path(ruta_plantilla)
        self.carpeta_salida = Path(carpeta_salida)
        self.usar_timestamp = usar_timestamp
        self.apu_provider = apu_provider
        self.incluir_revision = incluir_revision
        # Regla de cotizacion: el PU NACE del SQL/catalogo de costeo (semilla),
        # pero una vez en el libro es de Excel: el Presupuesto queda ligado con
        # formula a la celda PU DE VENTA de la hoja APU, que hace VLOOKUP a
        # Insumos — editar Insumos mueve PU y total. El PU oficial del catalogo
        # queda como fila de referencia en la hoja APU y en Revision.
        # Con usar_pu_oficial=True se vuelve al modo estatico (PU congelado).
        self.usar_pu_oficial = usar_pu_oficial
        self.wb: Optional[openpyxl.Workbook] = None
        self.ruta_salida: Optional[Path] = None
        self.timestamp: str = ""

        # Mapas de APUs por clave
        self.apu_por_clave: dict[str, dict] = {}       # clave -> apu (matriz)
        self.apu_sheet: dict[str, str] = {}            # clave -> nombre de hoja
        self.apu_pu_cell: dict[str, str] = {}          # clave -> "F45" (celda PU)

        # IDs LOCALES de los insumos copiados de la base general.
        # La clave general (catálogo/matrices) no existe en la plantilla local
        # (p.ej. las cuadrillas 1A1P / 1T2E-U no están en la hoja Cuadrillas),
        # por lo que cada insumo recibe un ID del libro (INS-001, INS-002, ...)
        # y TODAS las referencias lo usan: Insumos, líneas de APU y Cuadrillas.
        self._id_local_insumo: dict[str, str] = {}     # clave general -> INS-xxx
        self._agg_insumos: dict[str, dict] = {}        # clave general -> agregado
        self._rango_cuadrillas = "Cuadrillas!$A$4:$I$200"

        # Límites de filas (se actualizan al escribir datos)
        self._fila_insumos_inicio = 5  # después de headers
        self._fila_presupuesto_inicio = 8  # después de headers

    # ------------------------------------------------------------------
    # Paso 1: Copiar plantilla
    # ------------------------------------------------------------------

    def copiar_plantilla(self) -> Path:
        """Copia la plantilla base a la carpeta de salida con timestamp opcional."""
        if not self.ruta_plantilla.exists():
            raise FileNotFoundError(f"Plantilla no encontrada: {self.ruta_plantilla}")

        self.timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M")

        if self.usar_timestamp:
            subcarpeta = self.carpeta_salida / self.timestamp
            subcarpeta.mkdir(parents=True, exist_ok=True)
            self.ruta_salida = subcarpeta / "presupuesto.xlsx"
        else:
            self.carpeta_salida.mkdir(parents=True, exist_ok=True)
            self.ruta_salida = self.carpeta_salida / "presupuesto.xlsx"

        shutil.copy2(self.ruta_plantilla, self.ruta_salida)
        self.wb = openpyxl.load_workbook(self.ruta_salida)
        return self.ruta_salida

    # ------------------------------------------------------------------
    # Paso 2: Variables
    # ------------------------------------------------------------------

    def escribir_variables(self, proyecto: dict):
        """Llena la hoja Variables con los datos del proyecto.

        Busca cada variable por su label y escribe el valor en la columna B.
        """
        ws = self._hoja("Variables")

        # Mapa de labels parciales → valores
        mapa = {
            "nombre del proyecto": proyecto.get("nombre", ""),
            "clave del proyecto": proyecto.get("clave", ""),
            "cliente": proyecto.get("cliente", ""),
            "fecha de presupuesto": proyecto.get("fecha", self.timestamp),
            "indirectos de campo": proyecto.get("indirectos_campo", 0.14),
            "indirectos de oficina": proyecto.get("indirectos_oficina", 0.05),
            "financiamiento": proyecto.get("financiamiento", 0.0135),
            "utilidad": proyecto.get("utilidad", 0.08),
            "iva": proyecto.get("iva", 0.16),
            "isn estatal": proyecto.get("isn", 0.03),
            "rendimiento default": proyecto.get("rendimiento", 1.0),
        }

        for row in ws.iter_rows(min_row=5, max_row=30, min_col=1, max_col=1):
            cell = row[0]
            if not cell.value:
                continue
            label = str(cell.value).lower()
            for clave_busqueda, valor in mapa.items():
                if clave_busqueda in label:
                    try:
                        escribir_celda(ws, cell.row, 2, valor,
                                       font=font_valor_editable, alignment=align_center,
                                       number_format="0.00%" if isinstance(valor, float) and valor < 1 else None)
                    except AttributeError:
                        # MergedCell — saltar
                        pass
                    break

    # ------------------------------------------------------------------
    # Paso 3: Insumos (tabla maestra)
    # ------------------------------------------------------------------

    def preparar_apus(self, resultados_matcheo: list[dict]):
        """Asocia cada concepto con su APU (matriz) y prepara nombres de hoja.
        
        La clave del mapa es el id_entrada (ej. CIM-001), no la clave_sugerida de la BD.
        El nombre de la hoja APU se deriva del id_entrada.
        """
        if self.apu_provider is None:
            return
        usados = {str(s).lower() for s in self.wb.sheetnames}
        for res in resultados_matcheo:
            if res.get("estado") not in ("confirmado", "dudoso", "confirmado_ia"):
                continue
            id_entrada = res.get("id_entrada", "")
            clave_bd = res.get("clave_sugerida")
            if not id_entrada or not clave_bd:
                continue
            try:
                apu = self.apu_provider.apu_de(clave_bd)
            except Exception:
                apu = None
            if not apu or not apu.get("insumos"):
                continue
            # Usar id_entrada como clave (ej. CIM-001), no la clave de la BD
            self.apu_por_clave[id_entrada] = apu
            nombre = self._nombre_hoja_apu(id_entrada, usados)
            usados.add(nombre.lower())
            self.apu_sheet[id_entrada] = nombre

    @staticmethod
    def _nombre_hoja_apu(clave: str, usados: set) -> str:
        base = re.sub(r"[\\/*?:\[\]]", "_", str(clave)).strip()
        base = re.sub(r"\s+", " ", base)
        base = base[:31] or "APU"
        nombre = base
        i = 2
        while nombre.lower() in usados:
            suf = f"_{i}"
            nombre = base[:31 - len(suf)] + suf
            i += 1
        return nombre

    def _meta_precio(self, resultado: dict) -> dict:
        """Trazabilidad del precio de un concepto.

        Devuelve: pu_oficial, pu_apu, origen_precio, diferencia y la fuente del
        desglose (canonica / matrices). Es la fila que se muestra en la hoja
        Revision para que la diferencia entre el PU publicado y el calculado
        por el APU quede a la vista.
        """
        vacio = {"pu_oficial": None, "pu_apu": None, "origen": "",
                 "diferencia": None, "fuente": ""}
        id_entrada = resultado.get("id_entrada", "")
        clave = resultado.get("clave_sugerida")

        apu = self.apu_por_clave.get(id_entrada)
        if apu:
            return {
                "pu_oficial": apu.get("pu_oficial"),
                "pu_apu": apu.get("pu_apu"),
                "origen": apu.get("origen_precio") or "",
                "diferencia": apu.get("diferencia_publicado_apu"),
                "fuente": apu.get("fuente") or "",
            }

        if self.apu_provider is not None and clave:
            try:
                c = self.apu_provider.concepto(clave)
            except Exception:
                c = None
            if c:
                return {
                    "pu_oficial": c.get("precio_unitario"),
                    "pu_apu": c.get("precio_unitario_apu"),
                    "origen": c.get("origen_precio") or "",
                    "diferencia": c.get("diferencia_publicado_apu"),
                    "fuente": "canonica" if c.get("apu_id") else "",
                }

        # Sin fuente consultable: el PU del catálogo es el oficial
        return {**vacio, "pu_oficial": resultado.get("pu_sugerido")}

    def _pu_oficial(self, id_entrada: str, resultado: dict) -> float:
        """PU que manda segun el modelo canonico (publicado > APU > manual)."""
        apu = self.apu_por_clave.get(id_entrada) or {}
        pu = apu.get("pu_oficial")
        if pu is None or float(pu) <= 0:
            pu = resultado.get("pu_sugerido") or 0
        return round(float(pu), 4)

    def _agregar_insumos(self, resultados_matcheo: list[dict]) -> dict[str, dict]:
        """Agrega los insumos de todos los APUs por clave de la fuente general.

        Devuelve: clave general -> {tipo, desc, unidad, costo, cant, es_cuadrilla}
        La cantidad total de la obra = S (cantidad del insumo en el APU x
        cantidad de obra).
        """
        agg: dict[str, dict] = {}
        for res in resultados_matcheo:
            if res.get("estado") not in ("confirmado", "dudoso", "confirmado_ia"):
                continue
            id_entrada = res.get("id_entrada", "")
            apu = self.apu_por_clave.get(id_entrada)
            if not apu:
                continue
            cantidad_obra = float(res.get("cantidad", 0) or 0)
            for ins in apu.get("insumos", []):
                k = ins.get("clave")
                if not k:
                    continue
                a = agg.get(k)
                if a is None:
                    es_cuadrilla = ins.get("tipo", "").lower() in ("cuadrilla", "mano_obra")
                    a = {"tipo": ins.get("tipo", ""), "desc": ins.get("descripcion", ""),
                         "unidad": ins.get("unidad", ""), "costo": float(ins.get("costo", 0) or 0),
                         "cant": 0.0, "es_cuadrilla": es_cuadrilla, "tasas": {}}
                    agg[k] = a
                a["cant"] += float(ins.get("cantidad", 0) or 0) * cantidad_obra
                costo = float(ins.get("costo", 0) or 0)
                if costo:
                    # Si la misma cuadrilla aparece con tasas distintas se conserva
                    # la mas frecuente (la de la base general).
                    clave_tasa = round(costo, 4)
                    a["tasas"][clave_tasa] = a["tasas"].get(clave_tasa, 0) + 1
                    if not a.get("costo"):
                        a["costo"] = costo
        return agg

    @staticmethod
    def _tasa_mayoritaria(tasas: dict) -> float:
        """Tasa (costo unitario) mas frecuente de un insumo.

        En caso de empate se toma la menor (conservadora para el presupuesto).
        """
        if not tasas:
            return 0.0
        return max(tasas.items(), key=lambda kv: (kv[1], -kv[0]))[0]

    def _asignar_ids_locales(self, agg: dict[str, dict]):
        """Asigna a cada insumo copiado un ID local del libro (INS-001, ...).

        El orden es estable (tipo, descripcion, clave general) para que el
        libro sea reproducible entre ejecuciones.
        """
        self._id_local_insumo = {}
        orden = sorted(
            agg.keys(),
            key=lambda k: (str(agg[k].get("tipo", "")), str(agg[k].get("desc", "")).upper(), str(k)),
        )
        for i, k in enumerate(orden, 1):
            self._id_local_insumo[k] = f"INS-{i:03d}"

    def escribir_cuadrillas(self, resultados_matcheo: list[dict]):
        """Agrega a la hoja Cuadrillas las cuadrillas usadas en la obra.

        La plantilla trae las cuadrillas base (CUAD001...), pero las matrices
        usan claves propias (p.ej. 1A1P, 1T2E-U) que NO existen en esa hoja:
        el VLOOKUP devolvia 0 y la mano de obra valia $0 (presupuesto
        subvaluado).

        Cada cuadrilla copiada de la base general recibe su ID local (INS-xxx)
        y se escribe en la hoja con el $/Jornal de la fuente, de modo que:
          - el VLOOKUP de Insumos!E resuelva, y
          - la liga sea local (el libro no depende de claves externas).
        """
        if "Cuadrillas" not in self.wb.sheetnames:
            return
        agg = self._agg_insumos or self._agregar_insumos(resultados_matcheo)
        self._agg_insumos = agg
        if not agg:
            return
        self._asignar_ids_locales(agg)

        ws = self.wb["Cuadrillas"]
        # Primera fila libre despues de las cuadrillas de la plantilla
        ultima = 4
        for r in range(5, ws.max_row + 1):
            if ws.cell(row=r, column=1).value:
                ultima = r
        fila = ultima + 1

        # Estilo de referencia: la primera cuadrilla de la plantilla
        estilos = None
        if ws.cell(row=5, column=1).value:
            estilos = [copy(ws.cell(row=5, column=c)._style) for c in range(1, 11)]

        cuadrillas = [(self._id_local_insumo[k], k, a)
                      for k, a in agg.items() if a.get("es_cuadrilla")]
        cuadrillas.sort(key=lambda t: t[0])

        for id_local, clave_general, a in cuadrillas:
            for c in range(1, 11):
                if estilos:
                    ws.cell(row=fila, column=c)._style = copy(estilos[c - 1])
            ws.cell(row=fila, column=1, value=id_local)
            ws.cell(row=fila, column=2, value=a.get("desc", ""))
            # Columnas C-H (composicion) se dejan vacias: la fuente general no
            # trae el desglose por categoria, solo la tasa consolidada.
            tasa = self._tasa_mayoritaria(a.get("tasas") or {}) or float(a.get("costo", 0) or 0)
            celda_i = ws.cell(row=fila, column=9, value=round(tasa, 4))
            celda_i.number_format = "$#,##0.00"
            ws.cell(row=fila, column=10,
                    value=f"{a.get('desc', '')} (origen: {clave_general})")
            fila += 1

        if cuadrillas:
            self._rango_cuadrillas = f"Cuadrillas!$A$4:$I${fila - 1}"

    def escribir_insumos(self, resultados_matcheo: list[dict]):
        """Crea la hoja Insumos: lista consolidada a nivel MATERIAL/insumo.

        Se agrega de todos los APUs (matrices): por cada insumo, la cantidad
        total de la obra = S (cantidad del insumo en el APU x cantidad de obra).

        Cada insumo copiado de la base general recibe un ID LOCAL (INS-001...),
        que es el que se usa en todas las referencias (lineas de las hojas APU
        y hoja Cuadrillas). La clave general se conserva en la columna I
        ("Clave origen") para trazabilidad.

        Para tipo=cuadrilla: el Costo unitario es VLOOKUP a la hoja Cuadrillas (col I).
        Para otros tipos: costo directo editable.

        Columnas:
        A: Id local | B: Tipo | C: Descripcion | D: Unidad
        E: Costo unitario | F: Cantidad total | G: Importe | H: Notas
        I: Clave origen (base general)
        """
        if "Insumos" in self.wb.sheetnames:
            del self.wb["Insumos"]
        ws = self.wb.create_sheet("Insumos")

        merge_y_escribir(ws, "A1:I1", "INSUMOS - Lista de materiales y recursos (explosion de APUs)",
                         font=self.dna.font_titulo(), fill=self.dna.fill_verde_icel(), alignment=self.dna.align_center())
        for c in range(1, 10):
            ws.cell(row=1, column=c).fill = self.dna.fill_verde_icel()

        merge_y_escribir(ws, "A2:I2",
                         "Ids locales (INS-xxx) generados para este libro; la clave de la base "
                         "general queda en la columna I. Edita el Costo unitario (col E). "
                         "Para Cuadrillas, edita el $/Jornal en la hoja Cuadrillas.",
                         font=self.dna.font_pequeno(italic=True), alignment=self.dna.align_left())

        headers = ["Id local", "Tipo", "Descripcion", "Unidad",
                   "Costo unitario", "Cantidad total", "Importe", "Notas", "Clave origen"]
        anchos = [12, 14, 55, 10, 16, 16, 18, 24, 16]
        for i, (h, w) in enumerate(zip(headers, anchos), 1):
            escribir_celda(ws, 4, i, h, font=self.dna.font_header_tabla_fallback(blanco=True), fill=self.dna.fill_verde_encabezado(),
                           alignment=self.dna.align_center_nowrap())
            ws.column_dimensions[col_letra(i)].width = w
        aplicar_borde(ws, 4, 1, 9)

        # Agregacion por clave de insumo de la fuente general
        agg = self._agg_insumos or self._agregar_insumos(resultados_matcheo)
        self._agg_insumos = agg
        self._asignar_ids_locales(agg)

        r = 5
        _lbl_tipo = {"material": "Material", "mano_obra": "Cuadrilla",
                      "equipo": "Equipo", "herramienta": "Herramienta",
                      "auxiliar": "Auxiliar", "maquinaria": "Maquinaria"}

        # Rango de la hoja Cuadrillas para VLOOKUP (col A=ID local, col I=$/Jornal)
        rango_cuadrillas = self._rango_cuadrillas

        # Orden estable por ID local
        for k in sorted(agg.keys(), key=lambda x: self._id_local_insumo.get(x, x)):
            a = agg[k]
            es_cuadrilla = a.get("es_cuadrilla", False)
            id_local = self._id_local_insumo.get(k, k)

            escribir_celda(ws, r, 1, id_local, font=font_normal, alignment=align_center)
            tipo_texto = _lbl_tipo.get(str(a.get("tipo", "material")).lower(),
                                      a.get("tipo", "material").capitalize())
            escribir_celda(ws, r, 2, tipo_texto,
                           font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 3, a["desc"], font=font_normal, alignment=align_left)
            escribir_celda(ws, r, 4, a["unidad"], font=font_normal, alignment=align_center)

            # Columna E: Costo unitario
            if es_cuadrilla:
                # Para cuadrillas: VLOOKUP a hoja Cuadrillas (col A=ID local, col 9=$/Jornal)
                formula_vlookup = f'=IFERROR(VLOOKUP(A{r},{rango_cuadrillas},9,FALSE),0)'
                escribir_celda(ws, r, 5, formula_vlookup, font=font_normal_rojo,
                               alignment=align_right, number_format="$#,##0.00")
            else:
                # Materiales, equipos, etc.: costo directo editable
                escribir_celda(ws, r, 5, round(a["costo"], 4), font=font_normal_rojo,
                               alignment=align_right, number_format="$#,##0.00")

            escribir_celda(ws, r, 6, round(a["cant"], 4), font=font_normal,
                           alignment=align_right, number_format="#,##0.0000")
            escribir_celda(ws, r, 7, f"=F{r}*E{r}", font=font_negrita,
                           alignment=align_right, number_format="$#,##0.00")
            escribir_celda(ws, r, 8, "Cuadrilla" if es_cuadrilla else "",
                           font=font_normal, alignment=align_left)
            escribir_celda(ws, r, 9, k, font=font_normal, alignment=align_center)
            aplicar_borde(ws, r, 1, 9)
            r += 1

        self._fila_insumos_fin = r - 1
        ws.freeze_panes = "A5"

    def escribir_revision(self, resultados_matcheo: list[dict]):
        """Crea la hoja Revision con los resultados del matcheo.

        El usuario puede editar esta hoja para:
        - Confirmar/rechazar matches
        - Marcar conceptos como 'desglosar APU' (S/N)
        """
        if "Revision" in self.wb.sheetnames:
            del self.wb["Revision"]
        ws = self.wb.create_sheet("Revision")

        tiene_ia = any(r.get("clave_ia") or r.get("motivo_ia") for r in resultados_matcheo)
        # Trazabilidad del precio: el modelo canonico define la prioridad
        # (publicado > APU > manual). Se muestra el PU oficial, el PU que
        # calcula el APU y la diferencia, para que no quede oculta.
        n_cols = 15 if tiene_ia else 10
        n_cols += 5

        merge_y_escribir(ws, f"A1:{col_letra(n_cols)}1", "REVISIÓN DE MATCHEO — Confirmar antes de construir",
                         font=font_titulo_blanco, fill=fill_azul_oscuro, alignment=align_center)
        for c in range(1, n_cols + 1):
            ws.cell(row=1, column=c).fill = fill_azul_oscuro

        merge_y_escribir(ws, f"A2:{col_letra(n_cols)}2",
                         "Edita 'Estado' (confirmado/dudoso/sin_match) y 'Desglosar APU' (S/N). Luego re-ejecuta. "
                         "Las columnas de PU muestran el precio oficial del catálogo y el que calcula el APU.",
                         font=font_italica, alignment=align_left)

        # Headers
        headers = ["Id entrada", "Concepto entrada", "Unidad", "Cantidad",
                   "Clave sugerida", "Descripción sugerida", "Score", "Método",
                   "Estado", "Desglosar APU",
                   "PU oficial", "PU APU", "Dif. %", "Origen precio", "Fuente desglose"]
        anchos = [12, 45, 8, 10, 14, 45, 8, 10, 12, 14,
                  14, 14, 9, 14, 16]
        if tiene_ia:
            headers += ["Clave IA", "Confianza IA", "Motivo IA"]
            anchos += [14, 12, 40]
        for i, (h, w) in enumerate(zip(headers, anchos), 1):
            escribir_celda(ws, 4, i, h, font=font_header_blanco, fill=fill_azul_medio,
                           alignment=align_center)
            ws.column_dimensions[col_letra(i)].width = w
        aplicar_borde(ws, 4, 1, n_cols)

        # Filas de datos
        r = 5
        for resultado in resultados_matcheo:
            escribir_celda(ws, r, 1, resultado.get("id_entrada", ""),
                           font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 2, resultado.get("concepto_entrada", ""),
                           font=font_normal, alignment=align_left)
            escribir_celda(ws, r, 3, resultado.get("unidad_entrada", ""),
                           font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 4, resultado.get("cantidad", 0),
                           font=font_normal, alignment=align_center, number_format="#,##0.00")
            escribir_celda(ws, r, 5, resultado.get("clave_sugerida", ""),
                           font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 6, resultado.get("descripcion_sugerida", ""),
                           font=font_normal, alignment=align_left)
            escribir_celda(ws, r, 7, resultado.get("score", 0),
                           font=font_normal, alignment=align_center, number_format="0.0")

            # Método
            metodo = resultado.get("metodo", "")
            escribir_celda(ws, r, 8, metodo, font=font_normal, alignment=align_center)

            # Estado (editable) — colorear según estado
            estado = resultado.get("estado", "sin_match")
            color_estado = {"confirmado": "C6EFCE", "confirmado_ia": "C6EFCE", "dudoso": "FFEB9C", "sin_match": "FFC7CE"}
            fill_estado = PatternFill(
                start_color=color_estado.get(estado, "FFFFFF"),
                end_color=color_estado.get(estado, "FFFFFF"),
                fill_type="solid"
            )
            escribir_celda(ws, r, 9, estado, font=font_negrita, fill=fill_estado,
                           alignment=align_center)

            # Desglosar APU (editable por el usuario)
            escribir_celda(ws, r, 10, "N", font=font_normal, alignment=align_center)

            # Trazabilidad del precio (modelo canonico de costos)
            meta = self._meta_precio(resultado)
            pu_of = meta.get("pu_oficial")
            pu_apu = meta.get("pu_apu")
            escribir_celda(ws, r, 11, round(float(pu_of), 4) if pu_of else "",
                           font=font_normal, alignment=align_right, number_format="$#,##0.00")
            escribir_celda(ws, r, 12, round(float(pu_apu), 4) if pu_apu else "",
                           font=font_normal, alignment=align_right, number_format="$#,##0.00")
            if pu_of and pu_apu and float(pu_of) > 0:
                dif_pct = (float(pu_apu) - float(pu_of)) / float(pu_of)
                fill_dif = None
                if abs(dif_pct) > 0.01:
                    fill_dif = PatternFill(start_color="FFF2CC", end_color="FFF2CC",
                                           fill_type="solid")
                celda = escribir_celda(ws, r, 13, dif_pct, font=font_normal,
                                       alignment=align_right, number_format="+0.0%;-0.0%;0.0%")
                if fill_dif:
                    celda.fill = fill_dif
            else:
                escribir_celda(ws, r, 13, "", font=font_normal, alignment=align_right)
            escribir_celda(ws, r, 14, meta.get("origen") or "",
                           font=font_normal, alignment=align_center)
            fuente = meta.get("fuente") or ""
            celda_fuente = escribir_celda(ws, r, 15, fuente,
                                          font=font_normal, alignment=align_center)
            if fuente and fuente != "canonica":
                # fuente no contractual: se marca para que se note
                celda_fuente.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC",
                                                fill_type="solid")

            if tiene_ia:
                escribir_celda(ws, r, 16, resultado.get("clave_ia") or "",
                               font=font_normal_rojo, alignment=align_center)
                escribir_celda(ws, r, 17, resultado.get("confianza_ia", ""),
                               font=font_normal, alignment=align_center, number_format="0.00")
                escribir_celda(ws, r, 18, resultado.get("motivo_ia", ""),
                               font=font_italica, alignment=align_left)

            aplicar_borde(ws, r, 1, n_cols)
            r += 1

        ws.freeze_panes = "A5"

    # ------------------------------------------------------------------
    # Paso 5: Hojas APU dinámicas
    # ------------------------------------------------------------------

    def crear_hojas_apu(self, resultados_matcheo: list[dict]):
        """Crea una hoja APU por cada concepto que tenga matriz de APU."""
        for id_entrada, apu in self.apu_por_clave.items():
            nombre = self.apu_sheet.get(id_entrada) or self._nombre_hoja_apu(id_entrada, set())
            if nombre in self.wb.sheetnames:
                del self.wb[nombre]
            ws = self.wb.create_sheet(nombre)
            # Pasar id_entrada para que aparezca en el encabezado de la hoja APU
            pu_cell = self._escribir_hoja_apu(ws, id_entrada, apu)
            self.apu_pu_cell[id_entrada] = pu_cell

    def _escribir_hoja_apu(self, ws, clave: str, apu: dict) -> str:
        """Escribe una hoja APU con los insumos reales de la matriz.

        Columnas: A Id local | B Insumo | C Unidad | D Cantidad | E Costo unit. |
                  F Importe | G Clave origen
        La columna A usa el ID LOCAL del insumo (INS-xxx), que es el mismo que
        tiene en la hoja Insumos: asi el VLOOKUP de E resuelve siempre dentro
        del libro. La clave de la base general queda en la columna G.
        Devuelve la referencia de la celda PU (ej. 'F38').
        """
        descripcion = apu.get("descripcion", "")
        unidad = apu.get("unidad", "")

        # ---- Encabezado ----
        ws.merge_cells("A1:B1")
        escribir_celda(ws, 1, 1, "PONTEM — Análisis de Precio Unitario",
                       font=font_titulo_blanco, fill=fill_azul_oscuro, alignment=align_left)
        escribir_celda(ws, 1, 3, clave, font=font_header_blanco,
                       fill=fill_azul_oscuro, alignment=align_center)
        merge_y_escribir(ws, "D1:E1", f"Concepto:\n{descripcion}",
                         font=Font(name="Calibri", size=9, bold=True, color="FFFFFF"),
                         fill=fill_azul_oscuro,
                         alignment=Alignment(horizontal="left", vertical="center", wrap_text=True))
        # Logo en la esquina superior derecha
        try:
            logo = self.ruta_plantilla.parent / "logo.png"
            if logo.exists():
                img = XLImage(str(logo))
                ratio = (img.height / img.width) if img.width else 1
                img.width = 70
                img.height = int(70 * ratio)
                ws.add_image(img, "F1")
        except Exception:
            pass
        ws.row_dimensions[1].height = 46

        escribir_celda(ws, 2, 1, "Unidad:", font=font_negrita, alignment=align_right)
        escribir_celda(ws, 2, 2, unidad, font=font_normal, alignment=align_left)
        escribir_celda(ws, 2, 5, "PU:", font=font_negrita, alignment=align_right)

        # ---- Headers de tabla ----
        fill_th = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        headers_apu = ["Id local", "Insumo", "Unidad", "Cantidad", "Costo unit.", "Importe",
                       "Clave origen"]
        for i, h in enumerate(headers_apu, 1):
            escribir_celda(ws, 4, i, h, font=font_header_negro, fill=fill_th, alignment=align_center)
        aplicar_borde(ws, 4, 1, 7)

        ord_secciones = ["material", "mano_obra", "equipo", "herramienta", "auxiliar", "maquinaria"]
        lbl = {"material": "Materiales", "mano_obra": "Mano de Obra", "equipo": "Equipo",
               "herramienta": "Herramienta", "auxiliar": "Auxiliares", "maquinaria": "Maquinaria"}

        # agrupar insumos por tipo
        por_tipo: dict[str, list] = {}
        for ins in apu.get("insumos", []):
            por_tipo.setdefault(ins.get("tipo", "material"), []).append(ins)

        rango_ins = f"Insumos!$A$5:$G${getattr(self, '_fila_insumos_fin', 5)}"
        fila = 5
        cd_parts = []
        for tipo in ord_secciones:
            items = por_tipo.get(tipo)
            if not items:
                continue
            # título de sección
            fill_sec = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
            merge_y_escribir(ws, f"B{fila}:D{fila}", lbl.get(tipo, tipo.title()),
                             font=font_seccion, fill=fill_sec, alignment=align_left)
            escribir_celda(ws, fila, 1, "", font=font_seccion, fill=fill_sec, alignment=align_center)
            escribir_celda(ws, fila, 5, "", font=font_seccion, fill=fill_sec, alignment=align_center)
            escribir_celda(ws, fila, 6, f"=SUM(F{fila+1}:F{fila+len(items)})",
                           font=font_seccion, fill=fill_sec,
                           alignment=align_right, number_format="$#,##0.00")
            escribir_celda(ws, fila, 7, "", font=font_seccion, fill=fill_sec, alignment=align_center)
            aplicar_borde(ws, fila, 1, 7)
            fila_sec = fila
            fila += 1
            for ins in items:
                iclave = ins.get("clave", "")
                id_local = self._id_local_insumo.get(iclave, iclave)
                escribir_celda(ws, fila, 1, id_local, font=font_normal, alignment=align_center)
                escribir_celda(ws, fila, 2, ins.get("descripcion", ""), font=font_normal, alignment=align_left)
                escribir_celda(ws, fila, 3, ins.get("unidad", ""), font=font_normal, alignment=align_center)
                escribir_celda(ws, fila, 4, round(float(ins.get("cantidad", 0) or 0), 6),
                               font=font_normal, alignment=align_right, number_format="#,##0.0000")
                # costo unitario desde Insumos (editable) - recálculo vivo
                costo_formula = f"=IFERROR(VLOOKUP(A{fila},{rango_ins},{5},FALSE),{float(ins.get('costo',0) or 0)})"
                escribir_celda(ws, fila, 5, costo_formula, font=font_normal,
                               alignment=align_right, number_format="$#,##0.00")
                escribir_celda(ws, fila, 6, f"=D{fila}*E{fila}", font=font_negrita,
                               alignment=align_right, number_format="$#,##0.00")
                escribir_celda(ws, fila, 7, iclave, font=font_normal, alignment=align_center)
                aplicar_borde(ws, fila, 1, 7)
                fila += 1
            cd_parts.append(f"F{fila_sec}")

        # ---- CD, indirectos, utilidad, PU ----
        cd_formula = "=" + "+".join(cd_parts) if cd_parts else "=0"
        merge_y_escribir(ws, f"C{fila}:E{fila}", "COSTO DIRECTO (CD)",
                         font=font_subtotal, fill=fill_gris_seccion, alignment=align_left)
        escribir_celda(ws, fila, 6, cd_formula, font=font_subtotal, fill=fill_gris_seccion,
                       alignment=align_right, number_format="$#,##0.00")
        cd_row = fila
        fila += 1

        merge_y_escribir(ws, f"C{fila}:D{fila}", "Indirectos (% CD)",
                         font=font_normal, alignment=align_left)
        escribir_celda(ws, fila, 5, "=Variables!B12+Variables!B13", font=font_normal_rojo,
                       alignment=align_center, number_format="0%")
        escribir_celda(ws, fila, 6, f"=F{cd_row}*E{fila}", font=font_normal,
                       alignment=align_right, number_format="$#,##0.00")
        ind_row = fila
        fila += 1

        merge_y_escribir(ws, f"C{fila}:D{fila}", "Utilidad (% CD+CI)",
                         font=font_normal, alignment=align_left)
        escribir_celda(ws, fila, 5, "=Variables!B15", font=font_normal_rojo,
                       alignment=align_center, number_format="0%")
        escribir_celda(ws, fila, 6, f"=(F{cd_row}+F{ind_row})*E{fila}", font=font_normal,
                       alignment=align_right, number_format="$#,##0.00")
        util_row = fila
        fila += 1

        merge_y_escribir(ws, f"C{fila}:E{fila}", "PU DE VENTA",
                         font=font_subtotal, fill=fill_pu_venta, alignment=align_left)
        escribir_celda(ws, fila, 6, f"=F{cd_row}+F{ind_row}+F{util_row}",
                       font=font_subtotal, fill=fill_pu_venta,
                       alignment=align_right, number_format="$#,##0.00")
        pu_row = fila
        escribir_celda(ws, 2, 6, f"=F{pu_row}", font=font_subtotal,
                       alignment=align_right, number_format="$#,##0.00")

        # ---- PU oficial del catalogo (regla del modelo canonico) ----
        # El modelo define la prioridad: publicado > APU > manual. Si el APU
        # recalculado difiere del publicado, ambos quedan a la vista.
        pu_oficial = apu.get("pu_oficial")
        if pu_oficial:
            fila += 1
            merge_y_escribir(ws, f"C{fila}:E{fila}", "PU OFICIAL (CATALOGO)",
                             font=font_normal, alignment=align_left)
            escribir_celda(ws, fila, 6, round(float(pu_oficial), 4),
                           font=font_normal, alignment=align_right,
                           number_format="$#,##0.00")
            origen = apu.get("origen_precio") or ""
            fuente = apu.get("fuente") or ""
            nota = f"origen: {origen}" if origen else ""
            if fuente and fuente != "canonica":
                nota += f" | desglose: {fuente} (fuente no contractual)"
            if nota:
                merge_y_escribir(ws, f"A{fila}:B{fila}", nota,
                                 font=font_italica, alignment=align_left)

        # ---- Anchos e impresión ----
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 50
        ws.column_dimensions["C"].width = 10
        ws.column_dimensions["D"].width = 12
        ws.column_dimensions["E"].width = 14
        ws.column_dimensions["F"].width = 18
        ws.column_dimensions["G"].width = 16
        ws.page_setup.orientation = "portrait"
        ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = 0.4
        ws.page_margins.right = 0.4
        ws.page_margins.top = 0.5
        ws.page_margins.bottom = 0.5

        return f"F{pu_row}"

    # ------------------------------------------------------------------
    # Paso 6: Presupuesto
    # ------------------------------------------------------------------

    def escribir_portada(self, proyecto: dict, logo_path: Optional[str] = None,
                         resultados_matcheo: Optional[list] = None):
        """Crea la hoja Portada (Propuesta Económica ICEL) con identidad visual corporativa.
        
        Estructura basada en caratula.xlsx + Pie de pagina.png:
        - Logo ICEL + datos fiscales (esquina superior)
        - Ciudad/Fecha + Clave propuesta
        - Propuesta + Att'n + Proyecto
        - Texto introductorio
        - Tabla dinámica: Partidas → Conceptos con fórmulas
        - Bloque financiero: Subtotal, IVA (16%), Total
        - Condiciones, Exclusiones, Notas, Tiempo de ejecución
        - Firma + Pie de página gráfico
        """
        if "Portada" in self.wb.sheetnames:
            del self.wb["Portada"]
        ws = self.wb.create_sheet("Portada", 0)

        # Hoja limpia (sin cuadrícula). La retícula del ADN (A:Q) NO se aplica
        # aquí todavía: esta Portada aún escribe en 6 columnas y los anchos de
        # la retícula la destrozan (columnas de 1.5 con descripciones largas).
        # PENDIENTE: reescribir esta hoja con la retícula A:Q del ADN y el
        # CELL_MAP (docs/ADN_ICEL.md), tomando como referencia la carátula
        # original (H:/Mi unidad/ICEL/PUs/CASA VENADO/caratula.xlsx).
        self.dna.configurar_hoja_limpia(ws)
        anchos_portada = {"A": 8, "B": 65, "C": 12, "D": 12, "E": 16, "F": 18}
        for col, ancho in anchos_portada.items():
            ws.column_dimensions[col].width = ancho

        # --- FASE 1: ENCABEZADO CORPORATIVO Y METADATOS ---
        r = 1
        
        # Logo ICEL en A1
        try:
            if logo_path and os.path.exists(logo_path):
                img = XLImage(logo_path)
                ratio = (img.height / img.width) if img.width else 1
                img.width = self.dna.LOGO_ANCHO
                img.height = int(self.dna.LOGO_ANCHO * ratio)
                ws.add_image(img, "A1")
        except Exception:
            pass

        # Datos fiscales ICEL alineados a la derecha (columnas E-F)
        # Usamos filas 1-5 para bloque corporativo
        datos_fiscales = [
            "Vía Lombardía 5701, Suite 409, Edificio Corporativo Distrito 1, Chihuahua, Chih.",
            "ICEL Contratistas S.A. de C.V. | RFC: ICO200226NP1"
        ]
        for i, linea in enumerate(datos_fiscales):
            merge_y_escribir(ws, f"E{r+i}:F{r+i}", linea,
                             font=self.dna.font_cuerpo_fallback(size=9),
                             alignment=self.dna.align_right())
        
        # Ciudad y Fecha (fila 7)
        r = 7
        ciudad_fecha = proyecto.get("ciudad_fecha", "Chihuahua, Chih., 22 de noviembre de 2026")
        escribir_celda(ws, r, 1, ciudad_fecha, font=self.dna.font_cuerpo_fallback(size=11))

        # Clave de propuesta (columna F, fila 7) - alineado derecha
        clave_propuesta = proyecto.get("clave", "AD01 - OMA MTY FASE III - R01")
        escribir_celda(ws, r, 6, clave_propuesta, font=self.dna.font_partida(), alignment=self.dna.align_right())

        # Propuesta (fila 10)
        r = 10
        escribir_celda(ws, r, 1, "Propuesta:", font=self.dna.font_subtitulo_fallback())
        merge_y_escribir(ws, f"B{r}:F{r}", proyecto.get("propuesta", "Propuesta Económica"),
                         font=self.dna.font_cuerpo_fallback(bold=True, size=11))

        # Att'n (fila 13)
        r = 13
        escribir_celda(ws, r, 1, "Att'n:", font=self.dna.font_subtitulo_fallback())
        escribir_celda(ws, r, 2, proyecto.get("cliente_nombre", ""), font=self.dna.font_cuerpo_fallback())
        r += 1
        escribir_celda(ws, r, 2, proyecto.get("cliente_cargo", ""), font=self.dna.font_cuerpo_fallback())
        r += 1
        # Proyecto en negrita
        escribir_celda(ws, r, 2, proyecto.get("proyecto_nombre", ""), font=self.dna.font_partida())

        # Texto introductorio (fila 17)
        r = 17
        texto_intro = proyecto.get("texto_intro", 
            "Por medio de la presente, ICEL Contratistas S.A. de C.V. presenta su propuesta económica "
            "para la ejecución de los trabajos descritos a continuación, de acuerdo con los alcances, "
            "especificaciones y condiciones establecidas.")
        merge_y_escribir(ws, f"B{r}:F{r}", texto_intro,
                         font=self.dna.font_cuerpo_fallback(size=10),
                         alignment=self.dna.align_left())

        # --- FASE 2: TABLA DE PRESUPUESTO DINÁMICA ---
        r = 20
        escribir_celda(ws, r, 1, "I.- Propuesta Económica:", font=self.dna.font_subtitulo_fallback())
        r += 1

        # Obtener totales por partida del presupuesto para crear la tabla dinámica
        # self._partidas contiene (nombre_partida, celda_subtotal)
        cd = getattr(self, "_pres_cd", None)
        ind = getattr(self, "_pres_ind", None)
        util = getattr(self, "_pres_util", None)
        sin_iva = getattr(self, "_pres_sin_iva", None)
        iva = getattr(self, "_pres_iva", None)
        gran = getattr(self, "_pres_gran", None)

        # Tabla de conceptos de la Portada.
        # Se toma de la ESTRUCTURA registrada al escribir la hoja Presupuesto
        # (partidas, subpartidas y conceptos). Antes se raspaba la hoja, lo que
        # arrastraba el pie de totales como si fueran conceptos, duplicaba cada
        # renglón y copiaba fórmulas relativas que apuntaban a otras filas.
        estructura = getattr(self, "_estructura_pres", None) or \
            self._estructura_desde_matcheo(resultados_matcheo or [])

        # Encabezados de tabla en Portada (Verde ICEL)
        headers_portada = ["Id", "Concepto", "Unidad", "Cantidad", "P.U.", "Importe"]
        for i, h in enumerate(headers_portada, 1):
            cell = ws.cell(row=r, column=i, value=h)
            cell.font = self.dna.font_header_tabla_fallback(blanco=True)
            cell.fill = self.dna.fill_verde_encabezado()
            cell.border = self.dna.border_header_tabla()
            cell.alignment = self.dna.align_center_nowrap()
        r += 1

        primera_fila_concepto = None
        fila_inicio_partida = None
        nombre_partida_actual = ""

        def _cerrar_partida():
            """Escribe el subtotal de la partida actual (suma de la Portada)."""
            nonlocal r, fila_inicio_partida, nombre_partida_actual
            if fila_inicio_partida is None or not nombre_partida_actual:
                return
            merge_y_escribir(ws, f"B{r}:E{r}", f"Subtotal {nombre_partida_actual}",
                             font=self.dna.font_cuerpo_fallback(bold=True),
                             alignment=self.dna.align_left())
            escribir_celda(ws, r, 6, f"=SUM(F{fila_inicio_partida}:F{r - 1})",
                           font=self.dna.font_moneda(bold=True),
                           alignment=self.dna.align_right(),
                           number_format="$#,##0.00")
            for col in range(1, 7):
                ws.cell(row=r, column=col).border = self.dna.border_celda()
            fila_inicio_partida = None
            nombre_partida_actual = ""
            r += 1

        for item in estructura:
            tipo = item.get("tipo")

            if tipo == "partida":
                _cerrar_partida()
                nombre_partida_actual = item.get("nombre", "") or ""
                merge_y_escribir(ws, f"A{r}:F{r}", f"PARTIDA: {nombre_partida_actual}",
                                 font=self.dna.font_partida(),
                                 alignment=self.dna.align_left())
                # ADN ICEL: filas agrupadoras de partida con fondo suave #F2F5F0
                for col in range(1, 7):
                    ws.cell(row=r, column=col).border = self.dna.border_partida()
                    ws.cell(row=r, column=col).fill = self.dna.fill_gris_partida()
                r += 1
                continue

            if tipo == "subpartida":
                merge_y_escribir(ws, f"A{r}:F{r}", f"SUBPARTIDA: {item.get('nombre', '')}",
                                 font=self.dna.font_cuerpo_fallback(bold=True),
                                 alignment=self.dna.align_left())
                for col in range(1, 7):
                    ws.cell(row=r, column=col).border = self.dna.border_celda()
                    ws.cell(row=r, column=col).fill = self.dna.fill_verde_claro()
                r += 1
                continue

            # --- Concepto ---
            if primera_fila_concepto is None:
                primera_fila_concepto = r

            escribir_celda(ws, r, 1, item.get("id") or "", font=self.dna.font_cuerpo_fallback(), alignment=self.dna.align_center())
            escribir_celda(ws, r, 2, item.get("descripcion") or "", font=self.dna.font_cuerpo_fallback(), alignment=self.dna.align_left())
            escribir_celda(ws, r, 3, item.get("unidad") or "", font=self.dna.font_cuerpo_fallback(), alignment=self.dna.align_center())
            escribir_celda(ws, r, 4, item.get("cantidad"), font=self.dna.font_cuerpo_fallback(), alignment=self.dna.align_center(), number_format="#,##0.00")

            # PU: referencia a la hoja APU (válida desde cualquier hoja) o valor del catálogo
            pu = item.get("pu")
            if pu is not None:
                escribir_celda(ws, r, 5, pu,
                               font=self.dna.font_moneda(),
                               alignment=self.dna.align_right(),
                               number_format="$#,##0.00")

            # Importe: se calcula DENTRO de la Portada. En la Portada el P.U.
            # está en la columna E (a diferencia del Presupuesto, donde está en
            # F), así que el importe es =D{r}*E{r}. Copiar "=D8*F8" aquí
            # produciría una referencia circular (F sería la propia celda).
            escribir_celda(ws, r, 6, f"=D{r}*E{r}",
                           font=self.dna.font_cuerpo_fallback(),
                           alignment=self.dna.align_right(),
                           number_format="$#,##0.00")

            for col in range(1, 7):
                ws.cell(row=r, column=col).border = self.dna.border_celda()

            # Ajustar altura según descripción
            descripcion = item.get("descripcion")
            if descripcion:
                lineas = max(1, (len(str(descripcion)) // 60) + 1)
                ws.row_dimensions[r].height = max(15, lineas * 12.5)

            if fila_inicio_partida is None:
                fila_inicio_partida = r
            r += 1

        # Subtotal de la última partida
        _cerrar_partida()

        # Si no hay conceptos, fallback a resumen por partidas
        if primera_fila_concepto is None:
            for nombre_part, celda_sub in getattr(self, "_partidas", []):
                if " | " in nombre_part:
                    num_part, titulo_part = nombre_part.split(" | ", 1)
                else:
                    num_part, titulo_part = "", nombre_part
                escribir_celda(ws, r, 1, num_part, font=self.dna.font_partida(), alignment=self.dna.align_center())
                merge_y_escribir(ws, f"B{r}:F{r}", titulo_part,
                                 font=self.dna.font_partida(), alignment=self.dna.align_left())
                for col in range(1, 7):
                    ws.cell(row=r, column=col).border = self.dna.border_partida()
                r += 1
                escribir_celda(ws, r, 1, "", font=self.dna.font_cuerpo_fallback(), alignment=self.dna.align_center())
                merge_y_escribir(ws, f"B{r}:E{r}", "Subtotal " + nombre_part,
                                 font=self.dna.font_cuerpo_fallback(bold=True), alignment=self.dna.align_left())
                escribir_celda(ws, r, 6, f"=Presupuesto!{celda_sub}",
                               font=self.dna.font_moneda(bold=True), alignment=self.dna.align_right(),
                               number_format="$#,##0.00")
                for col in range(1, 7):
                    ws.cell(row=r, column=col).border = self.dna.border_celda()
                r += 1

        # --- FASE 3: TOTALES FINANCIEROS ---
        r += 1  # espacio en blanco

        def _fila_total(etiqueta: str, celda_presupuesto: Optional[str], *,
                        negrita: bool = False, con_fill: bool = False) -> int:
            """Escribe una fila del bloque financiero de la Portada.

            Todos los importes se referencian al Presupuesto con el prefijo de
            hoja ('=Presupuesto!G16'), de modo que no dependan de filas locales.
            """
            nonlocal r
            merge_y_escribir(ws, f"A{r}:E{r}", etiqueta,
                             font=self.dna.font_subtitulo_fallback() if not negrita
                             else self.dna.font_partida(size=12),
                             alignment=self.dna.align_right())
            if celda_presupuesto:
                escribir_celda(ws, r, 6, f"=Presupuesto!{celda_presupuesto}",
                               font=self.dna.font_moneda(bold=negrita),
                               alignment=self.dna.align_right(),
                               number_format="$#,##0.00")
            for col in range(1, 7):
                ws.cell(row=r, column=col).border = self.dna.border_celda()
                if con_fill:
                    ws.cell(row=r, column=col).fill = self.dna.fill_verde_claro()
            fila = r
            r += 1
            return fila

        # Costo directo
        _fila_total("Costo Directo", cd)
        # Indirectos
        _fila_total("Indirectos", ind)
        # Utilidad
        _fila_total("Utilidad", util)
        # Subtotal sin IVA
        fila_subtotal = _fila_total("Subtotal", sin_iva, negrita=True)
        # I.V.A. (16%)
        if iva:
            fila_iva = _fila_total("I.V.A. (16%)", iva)
        else:
            fila_iva = r
            merge_y_escribir(ws, f"A{r}:E{r}", "I.V.A. (16%)",
                             font=self.dna.font_subtitulo_fallback(),
                             alignment=self.dna.align_right())
            escribir_celda(ws, r, 6, f"=F{fila_subtotal}*0.16",
                           font=self.dna.font_moneda(),
                           alignment=self.dna.align_right(),
                           number_format="$#,##0.00")
            for col in range(1, 7):
                ws.cell(row=r, column=col).border = self.dna.border_celda()
            r += 1
        # Total (negrita, resaltado)
        if gran:
            _fila_total("Total", gran, negrita=True, con_fill=True)
        else:
            merge_y_escribir(ws, f"A{r}:E{r}", "Total",
                             font=self.dna.font_partida(size=12),
                             alignment=self.dna.align_right())
            escribir_celda(ws, r, 6, f"=F{fila_subtotal}+F{fila_iva}",
                           font=self.dna.font_moneda(bold=True),
                           alignment=self.dna.align_right(),
                           number_format="$#,##0.00")
            for col in range(1, 7):
                ws.cell(row=r, column=col).border = self.dna.border_celda()
                ws.cell(row=r, column=col).fill = self.dna.fill_verde_claro()
            r += 1
        r += 1

        # --- FASE 4: CONDICIONES, EXCLUSIONES Y NOTAS ---
        secciones_notas = proyecto.get("secciones_notas", [
            {
                "titulo": "II.- Condiciones Comerciales",
                "incisos": [
                    ("2.1", "Precios en pesos mexicanos (MXN)."),
                    ("2.2", "Vigencia de la propuesta: 30 días a partir de la fecha de emisión."),
                    ("2.3", "Los precios incluyen IVA (16%)."),
                    ("2.4", "Modelo de contrato a precio alzado; volúmenes de referencia."),
                    ("2.5", "Pago: 30% anticipo, 40% contra entrega de avances, 30% contra recepción final."),
                ]
            },
            {
                "titulo": "III.- Tiempo de Ejecución",
                "incisos": [
                    ("3.1", "Plazo estimado: 120 días naturales a partir de la orden de inicio."),
                    ("3.2", "El cronograma detallado se presentará en la junta de arranque."),
                ]
            },
            {
                "titulo": "IV.- Exclusiones",
                "incisos": [
                    ("4.1", "Permisos, licencias y trámites ante autoridades."),
                    ("4.2", "Estudios de mecánica de suelos, topografía y geología."),
                    ("4.3", "Servicios temporales (agua, luz, drenaje) durante la obra."),
                    ("4.4", "Mobiliario, equipo de computo y decoración no especificados."),
                ]
            },
            {
                "titulo": "V.- Notas Generales",
                "incisos": [
                    ("5.1", "Cualquier trabajo no contemplado se cotizará por separado."),
                    ("5.2", "Los precios son firmes por el plazo de vigencia de la propuesta."),
                    ("5.3", "Esta propuesta no constituye oferta legal vinculante hasta la firma del contrato."),
                ]
            },
        ])

        for seccion in secciones_notas:
            # Título de sección
            escribir_celda(ws, r, 1, seccion["titulo"], font=self.dna.font_subtitulo_fallback())
            r += 1
            
            for numero, texto in seccion["incisos"]:
                escribir_celda(ws, r, 1, numero, font=self.dna.font_cuerpo_fallback(size=10), alignment=self.dna.align_center())
                merge_y_escribir(ws, f"B{r}:F{r}", texto,
                                 font=self.dna.font_cuerpo_fallback(size=10),
                                 alignment=self.dna.align_left())
                # Ajustar altura de fila según longitud del texto
                lineas_estimadas = max(1, (len(texto) // 65) + 1)
                ws.row_dimensions[r].height = max(20, lineas_estimadas * 15)
                for col in range(1, 7):
                    ws.cell(row=r, column=col).border = self.dna.border_celda()
                r += 1
            
            r += 1  # espacio entre secciones

        # --- FASE 5: FIRMA Y PIE DE PÁGINA GRÁFICO ---
        r += 2
        firmante = proyecto.get("firmante", "Ing. Noé J. Sáenz Chávez")
        cargo_firmante = proyecto.get("cargo_firmante", "Director Operativo")
        
        merge_y_escribir(ws, f"A{r}:F{r}", firmante,
                         font=self.dna.font_partida(),
                         alignment=self.dna.align_center())
        r += 1
        merge_y_escribir(ws, f"A{r}:F{r}", cargo_firmante,
                         font=self.dna.font_subtitulo_fallback(),
                         alignment=self.dna.align_center())
        r += 2

        # Insertar Pie de pagina.png ajustado al ancho de la tabla (columnas A-F)
        try:
            footer_path = self.dna.FOOTER_PATH
            if os.path.exists(footer_path):
                img_footer = XLImage(footer_path)
                # Escalar al ancho aproximado de 6 columnas (~1300px)
                img_footer.width = 750
                img_footer.height = self.dna.FOOTER_ALTO
                ws.add_image(img_footer, f"A{r}")
        except Exception:
            pass

        # --- FASE 6: CONFIGURACIÓN DE PÁGINA ---
        self.dna.configurar_impresion(ws)

    def escribir_presupuesto(self, resultados_matcheo: list[dict], proyecto: dict):
        """Llena la hoja Presupuesto con formato tipo catálogo (mockup OMA).

        Estructura:
        - Encabezado: empresa, proyecto, cliente, fecha, cotización, revisión.
        - Tabla WBS: No. | Clave/WBS | Descripción | Cantidad | Unidad | PU | Importe | %
        - Partidas (azul), subpartidas (azul claro) con subtotales y % del total.
        - Conceptos con fórmulas VLOOKUP → Insumos.
        - Pie: costo directo, indirectos, utilidad, IVA, gran total.

        Columnas: A=No. | B=Clave/WBS | C=Descripción | D=Cantidad | E=Unidad
                  F=PU | G=Importe | H=% del total
        """
        if "Presupuesto" in self.wb.sheetnames:
            del self.wb["Presupuesto"]
        ws = self.wb.create_sheet("Presupuesto", 0)  # primera hoja
        self._partidas = []
        # Estructura real escrita (partidas, subpartidas y conceptos) para que la
        # Portada copie exactamente lo mismo, sin raspar el pie de totales.
        self._estructura_pres = []

        # ---- Paleta ----
        AZUL = "0070C0"
        AZUL_CLARO = "BDD7EE"
        AZUL_MUY_CLARO = "DDEBF7"
        GRIS_HEADER = "404040"
        VERDE_TOTAL = "E2EFDA"

        fill_partida = PatternFill(start_color=AZUL, end_color=AZUL, fill_type="solid")
        fill_subpartida = PatternFill(start_color=AZUL_CLARO, end_color=AZUL_CLARO, fill_type="solid")
        fill_concepto_banda = PatternFill(start_color=AZUL_MUY_CLARO, end_color=AZUL_MUY_CLARO, fill_type="solid")

        font_blanco_bold = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        font_negro_bold = Font(name="Arial", size=10, bold=True, color="000000")
        font_col_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")

        # ============================================================
        # ENCABEZADO (filas 1-5)
        # ============================================================
        nombre_proy = proyecto.get("nombre", "PROYECTO")
        cliente = proyecto.get("cliente", "")
        fecha = proyecto.get("fecha", self.timestamp)
        cotizacion = proyecto.get("cotizacion", "COT-001")
        revision = proyecto.get("revision", "R01")

        # Fila 1: banda de título
        merge_y_escribir(ws, "A1:H1", "PRESUPUESTO DE OBRA",
                         font=Font(name="Arial", size=14, bold=True, color="FFFFFF"),
                         fill=PatternFill(start_color=GRIS_HEADER, end_color=GRIS_HEADER, fill_type="solid"),
                         alignment=align_center)
        for c in range(1, 9):
            ws.cell(row=1, column=c).fill = PatternFill(start_color=GRIS_HEADER, end_color=GRIS_HEADER, fill_type="solid")
        ws.row_dimensions[1].height = 24

        # Fila 2: nombre del proyecto
        merge_y_escribir(ws, "A2:H2", nombre_proy,
                         font=Font(name="Arial", size=12, bold=True, color="000000"),
                         fill=PatternFill(start_color=VERDE_TOTAL, end_color=VERDE_TOTAL, fill_type="solid"),
                         alignment=align_left)
        for c in range(1, 9):
            ws.cell(row=2, column=c).fill = PatternFill(start_color=VERDE_TOTAL, end_color=VERDE_TOTAL, fill_type="solid")
        ws.row_dimensions[2].height = 20

        # Fila 3: cliente | fecha
        merge_y_escribir(ws, "A3:D3", f"Cliente: {cliente}",
                         font=font_negro_bold, alignment=align_left)
        merge_y_escribir(ws, "E3:H3", f"Fecha: {fecha}",
                         font=font_negro_bold, alignment=align_right)

        # Fila 4: cotización | revisión
        merge_y_escribir(ws, "A4:D4", f"Cotización No.: {cotizacion}",
                         font=font_negro_bold, alignment=align_left)
        merge_y_escribir(ws, "E4:H4", f"Revisión: {revision}",
                         font=font_negro_bold, alignment=align_right)

        # Fila 5: spacer
        ws.row_dimensions[5].height = 6

        # ============================================================
        # TABLA — HEADERS (fila 6)
        # ============================================================
        fila_header = 6
        headers = ["Id", "Clave PONTEM", "Descripción", "Cantidad", "Unidad", "P.U.", "Importe", "%"]
        anchos = [7, 11, 60, 11, 10, 15, 17, 9]
        for i, (h, w) in enumerate(zip(headers, anchos), 1):
            escribir_celda(ws, fila_header, i, h, font=font_col_header,
                           fill=PatternFill(start_color=GRIS_HEADER, end_color=GRIS_HEADER, fill_type="solid"),
                           alignment=align_center)
            ws.column_dimensions[col_letra(i)].width = w
        aplicar_borde(ws, fila_header, 1, 8)
        ws.row_dimensions[fila_header].height = 20
        # Ocultar columna B (Clave PONTEM) — solo referencia interna
        ws.column_dimensions["B"].hidden = True

        # ---- Datos ----
        # Conceptos con precio (confirmado/dudoso con clave)
        con_pu = [res for res in resultados_matcheo
                  if res["estado"] in ("confirmado", "dudoso", "confirmado_ia") and res.get("clave_sugerida")]
        rango_insumos = f"A5:J{self._fila_insumos_fin}" if hasattr(self, '_fila_insumos_fin') else "A:J"

        # Se incluyen TODOS los conceptos en el Presupuesto; los sin match se
        # muestran con PU 0 y marcados, para que el usuario los complete.
        todos = list(resultados_matcheo)

        # Agrupar por partida preservando orden
        grupos: dict[str, list[dict]] = {}
        orden_partidas: list[str] = []
        for res in todos:
            p = res.get("partida", "GENERAL") or "GENERAL"
            if p not in grupos:
                grupos[p] = []
                orden_partidas.append(p)
            grupos[p].append(res)

        # ---- Guardia de outliers: PU anómalo en líneas dudosas ----
        # Calcula la mediana de PU por unidad y marca como "revisar" las líneas
        # dudosas cuyo PU supere 12x esa mediana (evita totales absurdos).
        import statistics
        from src.matcher import normalizar_unidad as _norm_uni
        _pu_por_unidad: dict[str, list[float]] = {}
        for res in con_pu:
            pu = res.get("pu_sugerido") or 0
            if pu > 0:
                u = _norm_uni(res.get("unidad_entrada", ""))
                _pu_por_unidad.setdefault(u, []).append(pu)
        _mediana = {u: statistics.median(v) for u, v in _pu_por_unidad.items() if v}

        def _es_revisar(res) -> bool:
            if res.get("estado") in ("confirmado", "confirmado_ia"):
                return False
            id_entrada = res.get("id_entrada", "")
            u = _norm_uni(res.get("unidad_entrada", ""))
            med = _mediana.get(u, 0)

            # PU estático cotizado (modo oficial) o sugerido
            pu = res.get("pu_sugerido") or 0
            if med > 0 and pu > 12 * med:
                return True

            # Modo fórmula: el PU cotizado será el de la hoja APU, que puede
            # venir de la matriz (fuente no contractual). Si ese desglose da un
            # PU absurdo frente a la mediana de su unidad, se marca para
            # revisar aunque el PU oficial del catálogo parezca normal
            # (caso K78.7: oficial $1,369 vs matriz $826,546).
            apu = self.apu_por_clave.get(id_entrada)
            if apu:
                pu_desglose = float(apu.get("pu_matriz") or apu.get("pu_apu") or 0)
                if med > 0 and pu_desglose > 12 * med:
                    return True
            return False

        r = fila_header + 1
        num_concepto = 0
        filas_subtotales_partida: list[int] = []
        filas_porcentaje: list[tuple[int, bool]] = []  # (fila, es_partida)

        for idx_partida, partida_nombre in enumerate(orden_partidas, start=1):
            conceptos_partida = grupos[partida_nombre]

            # ---- Encabezado de PARTIDA ----
            wbs_partida = str(idx_partida)
            escribir_celda(ws, r, 1, wbs_partida, font=font_blanco_bold, fill=fill_partida, alignment=align_center)
            escribir_celda(ws, r, 2, "", font=font_blanco_bold, fill=fill_partida, alignment=align_left)
            merge_y_escribir(ws, f"C{r}:F{r}", partida_nombre, font=font_blanco_bold,
                             fill=fill_partida, alignment=align_left)
            escribir_celda(ws, r, 7, None, font=font_blanco_bold, fill=fill_partida, alignment=align_right)
            escribir_celda(ws, r, 8, None, font=font_blanco_bold, fill=fill_partida, alignment=align_center, number_format="0.0%")
            for c in range(1, 9):
                ws.cell(row=r, column=c).fill = fill_partida
            aplicar_borde(ws, r, 1, 8)
            fila_partida_header = r
            filas_porcentaje.append((r, True))
            self._estructura_pres.append({
                "tipo": "partida", "fila": fila_partida_header,
                "wbs": wbs_partida, "nombre": partida_nombre,
            })
            r += 1

            # ¿Subpartidas?
            tiene_sub = any(res.get("subpartida") for res in conceptos_partida)

            if not tiene_sub:
                primer_concepto = r
                for res in conceptos_partida:
                    num_concepto += 1
                    fila_concepto = r
                    if res.get("clave_sugerida") and res.get("estado") in ("confirmado", "dudoso", "confirmado_ia"):
                        r = self._escribir_fila_concepto(ws, r, res, rango_insumos, num_concepto,
                                                         revisar=_es_revisar(res))
                    else:
                        r = self._escribir_fila_sin_match(ws, r, res, num_concepto)
                    self._registrar_concepto_pres(ws, fila_concepto)
                ultimo_concepto = r - 1
                # Subtotal partida en G e importe de la fila de partida
                escribir_celda(ws, fila_partida_header, 7,
                               f"=SUM(G{primer_concepto}:G{ultimo_concepto})",
                               font=font_blanco_bold, fill=fill_partida,
                               alignment=align_right, number_format="$#,##0.00")
                filas_subtotales_partida.append(fila_partida_header)
            else:
                grupos_sub: dict[str, list[dict]] = {}
                orden_sub: list[str] = []
                for res in conceptos_partida:
                    sp = res.get("subpartida") or "(sin subpartida)"
                    if sp not in grupos_sub:
                        grupos_sub[sp] = []
                        orden_sub.append(sp)
                    grupos_sub[sp].append(res)

                filas_sub_marcadores: list[int] = []
                for idx_sub, sub_nombre in enumerate(orden_sub, start=1):
                    wbs_sub = f"{idx_partida}.{idx_sub}"
                    # Encabezado de SUBPARTIDA
                    escribir_celda(ws, r, 1, wbs_sub, font=font_negro_bold, fill=fill_subpartida, alignment=align_center)
                    escribir_celda(ws, r, 2, "", font=font_negro_bold, fill=fill_subpartida, alignment=align_left)
                    merge_y_escribir(ws, f"C{r}:F{r}", sub_nombre, font=font_negro_bold,
                                     fill=fill_subpartida, alignment=align_left)
                    escribir_celda(ws, r, 7, None, font=font_negro_bold, fill=fill_subpartida, alignment=align_right)
                    escribir_celda(ws, r, 8, None, font=font_negro_bold, fill=fill_subpartida, alignment=align_center, number_format="0.0%")
                    for c in range(1, 9):
                        ws.cell(row=r, column=c).fill = fill_subpartida
                    aplicar_borde(ws, r, 1, 8)
                    fila_sub_header = r
                    filas_porcentaje.append((r, False))
                    self._estructura_pres.append({
                        "tipo": "subpartida", "fila": fila_sub_header,
                        "wbs": wbs_sub, "nombre": sub_nombre,
                    })
                    r += 1

                    primer_concepto = r
                    for res in grupos_sub[sub_nombre]:
                        num_concepto += 1
                        fila_concepto = r
                        if res.get("clave_sugerida") and res.get("estado") in ("confirmado", "dudoso", "confirmado_ia"):
                            r = self._escribir_fila_concepto(ws, r, res, rango_insumos, num_concepto,
                                                             revisar=_es_revisar(res))
                        else:
                            r = self._escribir_fila_sin_match(ws, r, res, num_concepto)
                        self._registrar_concepto_pres(ws, fila_concepto)
                    ultimo_concepto = r - 1

                    escribir_celda(ws, fila_sub_header, 7,
                                   f"=SUM(G{primer_concepto}:G{ultimo_concepto})",
                                   font=font_negro_bold, fill=fill_subpartida,
                                   alignment=align_right, number_format="$#,##0.00")
                    filas_sub_marcadores.append(fila_sub_header)
                    filas_subtotales_partida.append(fila_sub_header)

                # Subtotal partida = suma de subtotales de subpartidas
                escribir_celda(ws, fila_partida_header, 7,
                               "=" + "+".join(f"G{fr}" for fr in filas_sub_marcadores),
                               font=font_blanco_bold, fill=fill_partida,
                               alignment=align_right, number_format="$#,##0.00")

            # Registrar partida para la portada (nombre + celda de subtotal)
            self._partidas.append((partida_nombre, f"G{fila_partida_header}"))

            # Espacio entre partidas
            r += 1

        # ============================================================
        # PIE — TOTALES
        # ============================================================
        r += 1
        fila_cd = r
        merge_y_escribir(ws, f"A{r}:F{r}", "SUBTOTAL GENERAL (Costo Directo)",
                         font=Font(name="Arial", size=10, bold=True, color="000000"),
                         fill=PatternFill(start_color=VERDE_TOTAL, end_color=VERDE_TOTAL, fill_type="solid"),
                         alignment=align_right)
        cd_formula = "=" + "+".join(f"G{fr}" for fr in filas_subtotales_partida) if filas_subtotales_partida else "=0"
        escribir_celda(ws, r, 7, cd_formula,
                       font=Font(name="Arial", size=10, bold=True, color="000000"),
                       fill=PatternFill(start_color=VERDE_TOTAL, end_color=VERDE_TOTAL, fill_type="solid"),
                       alignment=align_right, number_format="$#,##0.00")
        for c in range(1, 9):
            ws.cell(row=r, column=c).fill = PatternFill(start_color=VERDE_TOTAL, end_color=VERDE_TOTAL, fill_type="solid")
        aplicar_borde(ws, r, 1, 8)
        r += 1

        # Indirectos
        merge_y_escribir(ws, f"A{r}:E{r}", "Indirectos (% del CD)", font=font_negro_bold, alignment=align_right)
        escribir_celda(ws, r, 6, "=Variables!B12+Variables!B13", font=font_normal_rojo, alignment=align_center, number_format="0%")
        escribir_celda(ws, r, 7, f"=G{fila_cd}*F{r}", font=font_negro_bold, alignment=align_right, number_format="$#,##0.00")
        fila_ind = r
        r += 1

        # Utilidad
        merge_y_escribir(ws, f"A{r}:E{r}", "Utilidad (% del CD + CI)", font=font_negro_bold, alignment=align_right)
        escribir_celda(ws, r, 6, "=Variables!B15", font=font_normal_rojo, alignment=align_center, number_format="0%")
        escribir_celda(ws, r, 7, f"=(G{fila_cd}+G{fila_ind})*F{r}", font=font_negro_bold, alignment=align_right, number_format="$#,##0.00")
        fila_util = r
        r += 1

        # Total sin IVA
        merge_y_escribir(ws, f"A{r}:F{r}", "TOTAL SIN IVA", font=font_negro_bold,
                         fill=PatternFill(start_color=VERDE_TOTAL, end_color=VERDE_TOTAL, fill_type="solid"),
                         alignment=align_right)
        escribir_celda(ws, r, 7, f"=G{fila_cd}+G{fila_ind}+G{fila_util}", font=font_negro_bold,
                       fill=PatternFill(start_color=VERDE_TOTAL, end_color=VERDE_TOTAL, fill_type="solid"),
                       alignment=align_right, number_format="$#,##0.00")
        for c in range(1, 9):
            ws.cell(row=r, column=c).fill = PatternFill(start_color=VERDE_TOTAL, end_color=VERDE_TOTAL, fill_type="solid")
        fila_sin_iva = r
        r += 1

        # IVA
        merge_y_escribir(ws, f"A{r}:E{r}", "IVA", font=font_negro_bold, alignment=align_right)
        escribir_celda(ws, r, 6, "=Variables!B18", font=font_normal_rojo, alignment=align_center, number_format="0%")
        escribir_celda(ws, r, 7, f"=G{fila_sin_iva}*F{r}", font=font_negro_bold, alignment=align_right, number_format="$#,##0.00")
        fila_iva = r
        r += 1

        # GRAN TOTAL
        merge_y_escribir(ws, f"A{r}:F{r}", "GRAN TOTAL (IVA INCLUIDO)",
                         font=Font(name="Arial", size=11, bold=True, color="FFFFFF"),
                         fill=fill_negro, alignment=align_right)
        escribir_celda(ws, r, 7, f"=G{fila_sin_iva}+G{fila_iva}",
                       font=Font(name="Arial", size=11, bold=True, color="FFFFFF"),
                       fill=fill_negro, alignment=align_right, number_format="$#,##0.00")
        for c in range(1, 9):
            ws.cell(row=r, column=c).fill = fill_negro

        # Guardar referencias de totales para la portada
        self._pres_cd = f"G{fila_cd}"
        self._pres_ind = f"G{fila_ind}"
        self._pres_util = f"G{fila_util}"
        self._pres_sin_iva = f"G{fila_sin_iva}"
        self._pres_iva = f"G{fila_iva}"
        self._pres_gran = f"G{r}"

        # ---- Segunda pasada: % del total (base = Costo Directo) ----
        for fr, es_partida in filas_porcentaje:
            f_font = font_blanco_bold if es_partida else font_negro_bold
            f_fill = fill_partida if es_partida else fill_subpartida
            escribir_celda(ws, fr, 8, f"=IF($G${fila_cd}=0,0,G{fr}/$G${fila_cd})",
                           font=f_font, fill=f_fill,
                           alignment=align_center, number_format="0.0%")

        ws.freeze_panes = f"A{fila_header + 1}"
        self._preparar_impresion(ws)

    def _registrar_concepto_pres(self, ws, fila: int):
        """Registra en `_estructura_pres` una fila de concepto del Presupuesto.

        La Portada usa esta estructura para copiar partidas/subpartidas/conceptos
        sin raspar la hoja (lo que arrastraba el pie de totales como si fueran
        conceptos) y sin duplicar renglones.
        """
        self._estructura_pres.append({
            "tipo": "concepto",
            "fila": fila,
            "id": ws.cell(row=fila, column=1).value,
            "descripcion": ws.cell(row=fila, column=3).value,
            "cantidad": ws.cell(row=fila, column=4).value,
            "unidad": ws.cell(row=fila, column=5).value,
            "pu": ws.cell(row=fila, column=6).value,
        })

    @staticmethod
    def _estructura_desde_matcheo(resultados_matcheo: list[dict]) -> list[dict]:
        """Estructura mínima (partidas + conceptos) si no se escribió el Presupuesto."""
        estructura: list[dict] = []
        grupos: dict[str, list[dict]] = {}
        orden: list[str] = []
        for res in resultados_matcheo:
            p = res.get("partida", "GENERAL") or "GENERAL"
            if p not in grupos:
                grupos[p] = []
                orden.append(p)
            grupos[p].append(res)
        for p in orden:
            estructura.append({"tipo": "partida", "fila": None, "nombre": p})
            for res in grupos[p]:
                estructura.append({
                    "tipo": "concepto",
                    "fila": None,
                    "id": res.get("id_entrada", ""),
                    "descripcion": res.get("descripcion_sugerida") or res.get("concepto_entrada", ""),
                    "cantidad": res.get("cantidad", 0),
                    "unidad": res.get("unidad_sugerida") or res.get("unidad_entrada", ""),
                    "pu": res.get("pu_sugerido", 0),
                })
        return estructura

    def _escribir_fila_concepto(self, ws, r: int, resultado: dict, rango_insumos: str,
                                num_concepto: int = 0, revisar: bool = False) -> int:
        """Escribe una fila de concepto. Columnas:
        A=No. | B=Clave | C=Descripción | D=Cantidad | E=Unidad | F=PU | G=Importe.

        El PU se toma de la hoja APU del concepto (si existe), si no del catálogo.
        La hoja APU se identifica por id_entrada (ej. CIM-001), no por clave_sugerida.
        """
        id_entrada = resultado.get("id_entrada", "")
        clave = resultado["clave_sugerida"]
        descripcion = resultado.get("descripcion_sugerida") or resultado.get("concepto_entrada", "")
        unidad = resultado.get("unidad_sugerida") or resultado.get("unidad_entrada", "")

        fill_revisar = PatternFill(start_color="FDE9E9", end_color="FDE9E9", fill_type="solid") if revisar else None

        escribir_celda(ws, r, 1, id_entrada, font=font_normal, alignment=align_center)
        escribir_celda(ws, r, 2, clave, font=font_normal, alignment=align_center)
        escribir_celda(ws, r, 3, descripcion, font=font_normal, alignment=align_left)

        escribir_celda(ws, r, 4, resultado.get("cantidad", 0),
                       font=font_normal, alignment=align_center, number_format="#,##0.00")

        escribir_celda(ws, r, 5, unidad, font=font_normal, alignment=align_center)

        # PU: por defecto, formula viva a la celda PU DE VENTA de la hoja APU
        # (VLOOKUP a Insumos: editar Insumos mueve el total). El PU del SQL es
        # solo la semilla. Con usar_pu_oficial=True se congela el oficial del
        # catalogo como valor estatico (modo opt-in).
        if revisar:
            escribir_celda(ws, r, 6, 0, font=font_normal_rojo,
                           alignment=align_right, number_format="$#,##0.00")
        elif self.usar_pu_oficial:
            escribir_celda(ws, r, 6, self._pu_oficial(id_entrada, resultado),
                           font=font_normal, alignment=align_right, number_format="$#,##0.00")
        elif id_entrada in self.apu_sheet and id_entrada in self.apu_pu_cell:
            hoja = self.apu_sheet[id_entrada]
            celda_pu = self.apu_pu_cell[id_entrada]
            escribir_celda(ws, r, 6, f"='{hoja}'!{celda_pu}",
                           font=font_normal, alignment=align_right, number_format="$#,##0.00")
        else:
            escribir_celda(ws, r, 6, float(resultado.get("pu_sugerido", 0) or 0),
                           font=font_normal, alignment=align_right, number_format="$#,##0.00")

        escribir_celda(ws, r, 7, f"=D{r}*F{r}",
                       font=font_negrita, alignment=align_right, number_format="$#,##0.00")

        aplicar_borde(ws, r, 1, 8)

        if fill_revisar is not None:
            for c in range(1, 9):
                ws.cell(row=r, column=c).fill = fill_revisar

        desc_len = len(descripcion or "")
        largo = max(desc_len, len(str(clave)) + 6)
        lineas = max(1, (largo // 70) + 1)
        ws.row_dimensions[r].height = max(15, lineas * 12.5)
        return r + 1

    def _escribir_fila_sin_match(self, ws, r: int, resultado: dict, num_concepto: int = 0) -> int:
        """Fila para un concepto sin match: muestra el texto de entrada, PU 0 y
        se marca en rojo para que el usuario lo complete en Revisión."""
        fill_revisar = PatternFill(start_color="FDE9E9", end_color="FDE9E9", fill_type="solid")
        desc = resultado.get("concepto_entrada", "")

        escribir_celda(ws, r, 1, resultado.get("id_entrada", ""), font=font_normal, alignment=align_center)
        escribir_celda(ws, r, 2, "-", font=font_normal_rojo, alignment=align_center)
        escribir_celda(ws, r, 3, desc, font=font_normal, alignment=align_left)
        escribir_celda(ws, r, 4, resultado.get("cantidad", 0),
                       font=font_normal, alignment=align_center, number_format="#,##0.00")
        escribir_celda(ws, r, 5, resultado.get("unidad_entrada", ""),
                       font=font_normal, alignment=align_center)
        escribir_celda(ws, r, 6, 0, font=font_normal_rojo,
                       alignment=align_right, number_format="$#,##0.00")
        escribir_celda(ws, r, 7, f"=D{r}*F{r}",
                       font=font_negrita, alignment=align_right, number_format="$#,##0.00")

        aplicar_borde(ws, r, 1, 8)
        for c in range(1, 9):
            ws.cell(row=r, column=c).fill = fill_revisar

        lineas = max(1, (len(desc) // 60) + 1)
        ws.row_dimensions[r].height = max(15, lineas * 12.5)
        return r + 1

    def _preparar_impresion(self, ws):
        """Configura la hoja para impresión: orientación, ajuste a ancho, márgenes."""
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        try:
            ws.print_options.horizontalCentered = True
        except Exception:
            pass
        ws.page_margins.left = 0.3
        ws.page_margins.right = 0.3
        ws.page_margins.top = 0.4
        ws.page_margins.bottom = 0.4
        # Repetir fila de encabezado de tabla en cada página
        ws.print_title_rows = "1:6"

    # ------------------------------------------------------------------
    # Paso 7: Explosión de insumos (BOM)
    # ------------------------------------------------------------------

    def escribir_explosion(self, resultados_matcheo: list[dict]):
        """Crea la hoja Explosion con la explosión de insumos (BOM).

        Para conceptos con PU cerrado, lista cada concepto y su cantidad total.
        Para conceptos con APU desglosado, consolida los insumos de cada APU.

        Columnas:
        A: Clave insumo | B: Descripción | C: Tipo | D: Unidad
        E: Cantidad total | F: Costo unitario | G: Importe total
        """
        if "Explosion" in self.wb.sheetnames:
            del self.wb["Explosion"]
        ws = self.wb.create_sheet("Explosion")

        merge_y_escribir(ws, "A1:G1", "EXPLOSIÓN DE INSUMOS (BOM) — Cantidades totales por insumo",
                         font=font_titulo_blanco, fill=fill_azul_oscuro, alignment=align_center)
        for c in range(1, 8):
            ws.cell(row=1, column=c).fill = fill_azul_oscuro

        merge_y_escribir(ws, "A2:G2",
                         "Consolidado de todos los insumos de la obra. "
                         "Para conceptos PU cerrado: cada concepto es un insumo. "
                         "Para APUs desglosados: suma de cantidades × consumo unitario.",
                         font=font_italica, alignment=align_left)

        # Headers
        headers = ["Clave", "Descripción", "Tipo", "Unidad", "Cantidad total",
                   "Costo unitario", "Importe total"]
        anchos = [14, 50, 16, 10, 16, 16, 18]
        for i, (h, w) in enumerate(zip(headers, anchos), 1):
            escribir_celda(ws, 4, i, h, font=font_header_blanco, fill=fill_azul_medio,
                           alignment=align_center)
            ws.column_dimensions[col_letra(i)].width = w
        aplicar_borde(ws, 4, 1, 7)

        # Agrupar conceptos por clave (sumar cantidades)
        from collections import defaultdict
        insumos_agrupados = defaultdict(lambda: {"desc": "", "tipo": "", "unidad": "",
                                                   "cantidad": 0, "pu": 0})

        for resultado in resultados_matcheo:
            if resultado["estado"] not in ("confirmado", "dudoso", "confirmado_ia"):
                continue
            if not resultado.get("clave_sugerida"):
                continue
            clave = resultado["clave_sugerida"]
            insumos_agrupados[clave]["desc"] = resultado.get("descripcion_sugerida", "")
            insumos_agrupados[clave]["tipo"] = "concepto_repo"
            insumos_agrupados[clave]["unidad"] = resultado.get("unidad_sugerida", "")
            insumos_agrupados[clave]["cantidad"] += resultado.get("cantidad", 0)
            insumos_agrupados[clave]["pu"] = resultado.get("pu_sugerido", 0)

        # Escribir filas
        r = 5
        for clave in sorted(insumos_agrupados.keys()):
            ins = insumos_agrupados[clave]
            escribir_celda(ws, r, 1, clave, font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 2, ins["desc"], font=font_normal, alignment=align_left)
            escribir_celda(ws, r, 3, ins["tipo"], font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 4, ins["unidad"], font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 5, ins["cantidad"],
                           font=font_normal, alignment=align_center, number_format="#,##0.00")
            escribir_celda(ws, r, 6, ins["pu"],
                           font=font_normal, alignment=align_right, number_format="$#,##0.00")
            # Importe = Cantidad × PU
            escribir_celda(ws, r, 7, f"=E{r}*F{r}",
                           font=font_negrita, alignment=align_right, number_format="$#,##0.00")
            aplicar_borde(ws, r, 1, 7)
            r += 1

        # Total
        fila_fin = r - 1
        r += 1
        merge_y_escribir(ws, f"A{r}:F{r}", "TOTAL EXPLOSIÓN (debe coincidir con Costo Directo del Presupuesto)",
                         font=font_subtotal, fill=fill_gris_seccion, alignment=align_right)
        escribir_celda(ws, r, 7, SUM(f"G5:G{fila_fin}"),
                       font=font_subtotal, fill=fill_gris_seccion,
                       alignment=align_right, number_format="$#,##0.00")

        ws.freeze_panes = "A5"

    # ------------------------------------------------------------------
    # Guardar
    # ------------------------------------------------------------------

    def guardar(self) -> Path:
        """Guarda el archivo y retorna la ruta."""
        if self.wb is None:
            raise RuntimeError("No hay workbook. Llama a copiar_plantilla() primero.")

        self.wb.save(self.ruta_salida)
        size_kb = os.path.getsize(self.ruta_salida) / 1024
        print(f"  Guardado: {self.ruta_salida} ({size_kb:.1f} KB)")
        return self.ruta_salida

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _hoja(self, nombre: str) -> openpyxl.worksheet.worksheet.Worksheet:
        """Obtiene una hoja por nombre. Lanza error si no existe."""
        if nombre not in self.wb.sheetnames:
            raise ValueError(f"Hoja '{nombre}' no encontrada. Hojas disponibles: {self.wb.sheetnames}")
        return self.wb[nombre]


# =============================================================================
# Función de conveniencia: construir todo de una vez
# =============================================================================

def construir_presupuesto(
    ruta_plantilla: str,
    proyecto: dict,
    resultados_matcheo: list[dict],
    carpeta_salida: str = "salidas",
    usar_timestamp: bool = False,
    apu_provider=None,
    incluir_revision: bool = False,
    validar: bool = True,
    recalcular: bool = True,
    usar_pu_oficial: bool = False,
    empresa: Optional[str] = None,
) -> Path:
    """Construye el presupuesto completo en un solo llamado.

    Orden:
      1. Copiar plantilla
      2. Variables
      3. Preparar APUs (matrices)
      4. Cuadrillas (filas locales de las cuadrillas usadas)
      5. Insumos (lista de materiales agregada de los APUs, con IDs locales)
      6. Hojas APU (una por concepto con matriz)
      7. Presupuesto (PU: formula viva a la hoja APU, o el oficial estatico)
      8. Portada
      9. Guardar
     10. Validar el libro (opcional; si hay errores se copia como
         presupuesto_con_errores.xlsx)

    Regla de cotizacion: el PU nace del SQL/catalogo de costeo (semilla), pero
    en el libro queda vivo: el Presupuesto referencia con formula la celda
    PU DE VENTA de la hoja APU (VLOOKUP a Insumos), de modo que editar
    Insumos en Excel mueve el PU y el total. Con `usar_pu_oficial=True` se
    congela el PU oficial del catalogo (v_concepto_pu.precio_unitario,
    publicado > APU > manual) como valor estatico.

    `empresa` selecciona el ADN visual (ej. 'icel'). None -> empresa por
    defecto del ADN. Puede venir en proyecto['empresa'].
    """
    empresa_efectiva = empresa or proyecto.get("empresa")
    c = Constructor(ruta_plantilla, carpeta_salida, usar_timestamp,
                    apu_provider=apu_provider, incluir_revision=incluir_revision,
                    usar_pu_oficial=usar_pu_oficial,
                    empresa=empresa_efectiva)

    print("[1/8] Copiando plantilla...")
    c.copiar_plantilla()

    # Limpieza: quitar Notas y ocultar hoja-motor de parámetros
    if "Notas" in c.wb.sheetnames:
        del c.wb["Notas"]
    for _h in ("Parametros",):
        if _h in c.wb.sheetnames:
            c.wb[_h].sheet_state = "hidden"

    print("[2/8] Escribiendo Variables...")
    c.escribir_variables(proyecto)

    print("[3/8] Preparando APUs (matrices)...")
    c.preparar_apus(resultados_matcheo)

    print("[4/8] Escribiendo Cuadrillas (IDs locales)...")
    c.escribir_cuadrillas(resultados_matcheo)

    print("[5/8] Escribiendo Insumos (materiales)...")
    c.escribir_insumos(resultados_matcheo)

    print("[6/8] Creando hojas APU...")
    c.crear_hojas_apu(resultados_matcheo)

    print("[7/8] Escribiendo Presupuesto...")
    c.escribir_presupuesto(resultados_matcheo, proyecto)

    print("[8/8] Escribiendo Portada...")
    logo = Path(ruta_plantilla).parent / "logo.png"
    c.escribir_portada(proyecto, str(logo) if logo.exists() else None, resultados_matcheo)

    if incluir_revision:
        c.escribir_revision(resultados_matcheo)

    print("\nGuardando...")
    ruta = c.guardar()

    # --- Validación de integridad del libro ---
    if validar:
        try:
            from src.validador import validar_libro, formatear_reporte
            print("\nValidando libro...")
            reporte = validar_libro(str(ruta), recalcular=recalcular)
            print(formatear_reporte(reporte))
            if reporte["errores"]:
                # Se conserva una copia con errores para depuración
                ruta_errores = ruta.with_name("presupuesto_con_errores.xlsx")
                try:
                    shutil.copy2(ruta, ruta_errores)
                    print(f"Copia con errores: {ruta_errores}")
                except Exception:
                    pass
        except Exception as e:
            print(f"[!] No se pudo validar el libro: {e}")

    return ruta