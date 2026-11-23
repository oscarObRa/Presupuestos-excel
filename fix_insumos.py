import re

with open('src/constructor.py', 'r', encoding='utf-8') as f:
    content = f.read()

old = '''def escribir_insumos(self, resultados_matcheo: list[dict], indice=None):
        """Crea la hoja Insumos: lista consolidada a nivel MATERIAL/insumo.

        Se agrega de todos los APUs (matrices): por cada insumo, la cantidad
        total de la obra = Σ (cantidad del insumo en el APU × cantidad de obra).

        Columnas:
        A: Clave | B: Tipo | C: Descripción | D: Unidad
        E: Costo unitario | F: Cantidad total | G: Importe | H: Notas
        """
        if "Insumos" in self.wb.sheetnames:
            del self.wb["Insumos"]
        ws = self.wb.create_sheet("Insumos")

        merge_y_escribir(ws, "A1:H1", "INSUMOS — Lista de materiales y recursos (explosión de APUs)",
                         font=font_titulo_blanco, fill=fill_azul_oscuro, alignment=align_center)
        for c in range(1, 9):
            ws.cell(row=1, column=c).fill = fill_azul_oscuro

        merge_y_escribir(ws, "A2:H2",
                         "Edita el Costo unitario (col E). Las hojas APU y el Presupuesto se recalculan.",
                         font=font_italica, alignment=align_left)

        headers = ["Clave", "Tipo", "Descripción", "Unidad",
                   "Costo unitario", "Cantidad total", "Importe", "Notas"]
        anchos = [14, 14, 55, 10, 16, 16, 18, 24]
        for i, (h, w) in enumerate(zip(headers, anchos), 1):
            escribir_celda(ws, 4, i, h, font=font_header_blanco, fill=fill_azul_medio,
                           alignment=align_center)
            ws.column_dimensions[col_letra(i)].width = w
        aplicar_borde(ws, 4, 1, 8)

        # Agregación por clave de insumo
        agg: dict[str, dict] = {}
        for res in resultados_matcheo:
            if res.get("estado") not in ("confirmado", "dudoso", "confirmado_ia"):
                continue
            clave = res.get("clave_sugerida")
            apu = self.apu_por_clave.get(clave)
            if not apu:
                continue
            cantidad_obra = float(res.get("cantidad", 0) or 0)
            for ins in apu.get("insumos", []):
                k = ins.get("clave")
                if not k:
                    continue
                a = agg.get(k)
                if a is None:
                    a = {"tipo": ins.get("tipo", ""), "desc": ins.get("descripcion", ""),
                         "unidad": ins.get("unidad", ""), "costo": float(ins.get("costo", 0) or 0),
                         "cant": 0.0}
                    agg[k] = a
                a["cant"] += float(ins.get("cantidad", 0) or 0) * cantidad_obra

        r = 5
        _lbl_tipo = {"material": "Material", "mano_obra": "Cuadrilla",
                      "equipo": "Equipo", "herramienta": "Herramienta",
                      "auxiliar": "Auxiliar", "maquinaria": "Maquinaria"}
        for k in sorted(agg.keys()):
            a = agg[k]
            escribir_celda(ws, r, 1, k, font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 2, _lbl_tipo.get(a["tipo"], a["tipo"]),
                           font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 3, a["desc"], font=font_normal, alignment=align_left)
            escribir_celda(ws, r, 4, a["unidad"], font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 5, round(a["costo"], 4), font=font_normal_rojo,
                           alignment=align_right, number_format="$#,##0.00")
            escribir_celda(ws, r, 6, round(a["cant"], 4), font=font_normal,
                           alignment=align_right, number_format="#,##0.0000")
            escribir_celda(ws, r, 7, f"=F{r}*E{r}", font=font_negrita,
                           alignment=align_right, number_format="$#,##0.00")
            escribir_celda(ws, r, 8, "", font=font_normal, alignment=align_left)
            aplicar_borde(ws, r, 1, 8)
            r += 1

        self._fila_insumos_fin = r - 1
        ws.freeze_panes = "A5"'''

new_code = '''def escribir_insumos(self, resultados_matcheo: list[dict], indice=None):
        """Crea la hoja Insumos: lista consolidada a nivel MATERIAL/insumo.

        Se agrega de todos los APUs (matrices): por cada insumo, la cantidad
        total de la obra = Σ (cantidad del insumo en el APU × cantidad de obra).

        Para tipo=cuadrilla: el Costo unitario es VLOOKUP a la hoja Cuadrillas (col C).
        Para otros tipos: costo directo editable.

        Columnas:
        A: Clave | B: Tipo | C: Descripción | D: Unidad
        E: Costo unitario | F: Cantidad total | G: Importe | H: Notas
        \"\"\"
        if "Insumos" in self.wb.sheetnames:
            del self.wb["Insumos"]
        ws = self.wb.create_sheet("Insumos")

        merge_y_escribir(ws, "A1:H1", "INSUMOS — Lista de materiales y recursos (explosión de APUs)",
                         font=DNA().font_titulo(), fill=DNA().fill_azul_oscuro(), alignment=DNA().align_center())
        for c in range(1, 9):
            ws.cell(row=1, column=c).fill = DNA().fill_azul_oscuro()

        merge_y_escribir(ws, "A2:H2",
                         "Edita el Costo unitario (col E). Para Cuadrillas, edita en hoja Cuadrillas (col C).",
                         font=DNA().font_italica(), alignment=DNA().align_left())

        headers = ["Clave", "Tipo", "Descripción", "Unidad",
                   "Costo unitario", "Cantidad total", "Importe", "Notas"]
        anchos = [14, 14, 55, 10, 16, 16, 18, 24]
        for i, (h, w) in enumerate(zip(headers, anchos), 1):
            escribir_celda(ws, 4, i, h, font=DNA().font_header(), fill=DNA().fill_header_azul(),
                           alignment=DNA().align_center())
            ws.column_dimensions[col_letra(i)].width = w
        aplicar_borde(ws, 4, 1, 8)

        # Agregación por clave de insumo
        agg: dict[str, dict] = {}
        for res in resultados_matcheo:
            if res.get("estado") not in ("confirmado", "dudoso", "confirmado_ia"):
                continue
            clave = res.get("clave_sugerida")
            apu = self.apu_por_clave.get(clave)
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
                         "cant": 0.0, "es_cuadrilla": es_cuadrilla}
                    agg[k] = a
                a["cant"] += float(ins.get("cantidad", 0) or 0) * cantidad_obra

        r = 5
        _lbl_tipo = {"material": "Material", "mano_obra": "Cuadrilla",
                      "equipo": "Equipo", "herramienta": "Herramienta",
                      "auxiliar": "Auxiliar", "maquinaria": "Maquinaria"}
        
        # Rango de la hoja Cuadrillas para VLOOKUP (col A=ID, col C=$/Jornal)
        rango_cuadrillas = "Cuadrillas!$A$4:$C$200"

        for k in sorted(agg.keys()):
            a = agg[k]
            es_cuadrilla = a.get("es_cuadrilla", False)
            
            escribir_celda(ws, r, 1, k, font=fuente_normal, alignment=align_center)
            escribir_celda(ws, r, 2, a.get("tipo", "material").capitalize(),
                           font=fuente_normal, alignment=align_center)
            escribir_celda(ws, r, 3, a["desc"], font=fuente_normal, alignment=align_left)
            escribir_celda(ws, r, 4, a["unidad"], font=fuente_normal, alignment=align_center)
            
            # Columna E: Costo unitario
            if a.get("es_cuadrilla", False):
                # Para cuadrillas: VLOOKUP a hoja Cuadrillas (col A=ID, col C=$/Jornal)
                formula_vlookup = f'=IFERROR(VLOOKUP(A{r},Cuadrillas!$A$4:$C$200,3,FALSE),0)'
                escribir_celda(ws, r, 5, formula_vlookup, font=fuente_normal_rojo,
                               alignment=align_right, number_format="$#,##0.00")
            else:
                # Materiales, equipos, etc.: costo directo editable
                escribir_celda(ws, r, 5, round(a["costo"], 4), font=fuente_normal_rojo,
                               alignment=align_right, number_format="$#,##0.00")
            
            escribir_celda(ws, r, 6, round(a["cant"], 4), font=fuente_normal,
                           alignment=align_right, number_format="#,##0.0000")
            escribir_celda(ws, r, 7, f"=F{r}*E{r}", font=fuente_negrita,
                           alignment=align_right, number_format="$#,##0.00")
            escribir_celda(ws, r, 8, "Cuadrilla" if a.get("es_cuadrilla") else "", 
                           font=fuente_normal, alignment=align_left)
            aplicar_borde(ws, r, 1, 8)
            r += 1

        self._fila_insumos_fin = r - 1
        ws.freeze_panes = "A5"'''

with open('src/constructor.py', 'r', encoding='utf-8') as f:
    content = f.read()

old = '''def escribir_insumos(self, resultados_matcheo: list[dict], indice=None):
        """Crea la hoja Insumos: lista consolidada a nivel MATERIAL/insumo.

        Se agrega de todos los APUs (matrices): por cada insumo, la cantidad
        total de la obra = Σ (cantidad del insumo en el APU × cantidad de obra).

        Columnas:
        A: Clave | B: Tipo | C: Descripción | D: Unidad
        E: Costo unitario | F: Cantidad total | G: Importe | H: Notas
        """
        if "Insumos" in self.wb.sheetnames:
            del self.wb["Insumos"]
        ws = self.wb.create_sheet("Insumos")

        merge_y_escribir(ws, "A1:H1", "INSUMOS — Lista de materiales y recursos (explosión de APUs)",
                         font=font_titulo_blanco, fill=fill_azul_oscuro, alignment=align_center)
        for c in range(1, 9):
            ws.cell(row=1, column=c).fill = fill_azul_oscuro

        merge_y_escribir(ws, "A2:H2",
                         "Edita el Costo unitario (col E). Las hojas APU y el Presupuesto se recalculan.",
                         font=font_italica, alignment=align_left)

        headers = ["Clave", "Tipo", "Descripción", "Unidad",
                   "Costo unitario", "Cantidad total", "Importe", "Notas"]
        anchos = [14, 14, 55, 10, 16, 16, 18, 24]
        for i, (h, w) in enumerate(zip(headers, anchos), 1):
            escribir_celda(ws, 4, i, h, font=font_header_blanco, fill=fill_azul_medio,
                           alignment=align_center)
            ws.column_dimensions[col_letra(i)].width = w
        aplicar_borde(ws, 4, 1, 8)

        # Agregación por clave de insumo
        agg: dict[str, dict] = {}
        for res in resultados_matcheo:
            if res.get("estado") not in ("confirmado", "dudoso", "confirmado_ia"):
                continue
            clave = res.get("clave_sugerida")
            apu = self.apu_por_clave.get(clave)
            if not apu:
                continue
            cantidad_obra = float(res.get("cantidad", 0) or 0)
            for ins in apu.get("insumos", []):
                k = ins.get("clave")
                if not k:
                    continue
                a = agg.get(k)
                if a is None:
                    a = {"tipo": ins.get("tipo", ""), "desc": ins.get("descripcion", ""),
                         "unidad": ins.get("unidad", ""), "costo": float(ins.get("costo", 0) or 0),
                         "cant": 0.0}
                    agg[k] = a
                a["cant"] += float(ins.get("cantidad", 0) or 0) * cantidad_obra

        r = 5
        _lbl_tipo = {"material": "Material", "mano_obra": "Cuadrilla",
                      "equipo": "Equipo", "herramienta": "Herramienta",
                      "auxiliar": "Auxiliar", "maquinaria": "Maquinaria"}
        for k in sorted(agg.keys()):
            a = agg[k]
            escribir_celda(ws, r, 1, k, font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 2, _lbl_tipo.get(a["tipo"], a["tipo"]),
                           font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 3, a["desc"], font=font_normal, alignment=align_left)
            escribir_celda(ws, r, 4, a["unidad"], font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 5, round(a["costo"], 4), font=font_normal_rojo,
                           alignment=align_right, number_format="$#,##0.00")
            escribir_celda(ws, r, 6, round(a["cant"], 4), font=font_normal,
                           alignment=align_right, number_format="#,##0.0000")
            escribir_celda(ws, r, 7, f"=F{r}*E{r}", font=font_negrita,
                           alignment=align_right, number_format="$#,##0.00")
            escribir_celda(ws, r, 8, "", font=font_normal, alignment=align_left)
            aplicar_borde(ws, r, 1, 8)
            r += 1

        self._fila_insumos_fin = r - 1
        ws.freeze_panes = "A5"'''

with open('src/constructor.py', 'r', encoding='utf-8') as f:
    content = f.read()

old = '''def escribir_insumos(self, resultados_matcheo: list[dict], indice=None):
        """Crea la hoja Insumos: lista consolidada a nivel MATERIAL/insumo.

        Se agrega de todos los APUs (matrices): por cada insumo, la cantidad
        total de la obra = Σ (cantidad del insumo en el APU × cantidad de obra).

        Columnas:
        A: Clave | B: Tipo | C: Descripción | D: Unidad
        E: Costo unitario | F: Cantidad total | G: Importe | H: Notas
        """
        if "Insumos" in self.wb.sheetnames:
            del self.wb["Insumos"]
        ws = self.wb.create_sheet("Insumos")

        merge_y_escribir(ws, "A1:H1", "INSUMOS — Lista de materiales y recursos (explosión de APUs)",
                         font=font_titulo_blanco, fill=fill_azul_oscuro, alignment=align_center)
        for c in range(1, 9):
            ws.cell(row=1, column=c).fill = fill_azul_oscuro

        merge_y_escribir(ws, "A2:H2",
                         "Edita el Costo unitario (col E). Las hojas APU y el Presupuesto se recalculan.",
                         font=font_italica, alignment=align_left)

        headers = ["Clave", "Tipo", "Descripción", "Unidad",
                   "Costo unitario", "Cantidad total", "Importe", "Notas"]
        anchos = [14, 14, 55, 10, 16, 16, 18, 24]
        for i, (h, w) in enumerate(zip(headers, anchos), 1):
            escribir_celda(ws, 4, i, h, font=font_header_blanco, fill=fill_azul_medio,
                           alignment=align_center)
            ws.column_dimensions[col_letra(i)].width = w
        aplicar_borde(ws, 4, 1, 8)

        # Agregación por clave de insumo
        agg: dict[str, dict] = {}
        for res in resultados_matcheo:
            if res.get("estado") not in ("confirmado", "dudoso", "confirmado_ia"):
                continue
            clave = res.get("clave_sugerida")
            apu = self.apu_por_clave.get(clave)
            if not apu:
                continue
            cantidad_obra = float(res.get("cantidad", 0) or 0)
            for ins in apu.get("insumos", []):
                k = ins.get("clave")
                if not k:
                    continue
                a = agg.get(k)
                if a is None:
                    a = {"tipo": ins.get("tipo", ""), "desc": ins.get("descripcion", ""),
                         "unidad": ins.get("unidad", ""), "costo": float(ins.get("costo", 0) or 0),
                         "cant": 0.0}
                    agg[k] = a
                a["cant"] += float(ins.get("cantidad", 0) or 0) * cantidad_obra

        r = 5
        _lbl_tipo = {"material": "Material", "mano_obra": "Cuadrilla",
                      "equipo": "Equipo", "herramienta": "Herramienta",
                      "auxiliar": "Auxiliar", "maquinaria": "Maquinaria"}
        for k in sorted(agg.keys()):
            a = agg[k]
            escribir_celda(ws, r, 1, k, font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 2, _lbl_tipo.get(a["tipo"], a["tipo"]),
                           font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 3, a["desc"], font=font_normal, alignment=align_left)
            escribir_celda(ws, r, 4, a["unidad"], font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 5, round(a["costo"], 4), font=font_normal_rojo,
                           alignment=align_right, number_format="$#,##0.00")
            escribir_celda(ws, r, 6, round(a["cant"], 4), font=font_normal,
                           alignment=align_right, number_format="#,##0.0000")
            escribir_celda(ws, r, 7, f"=F{r}*E{r}", font=font_negrita,
                           alignment=align_right, number_format="$#,##0.00")
            escribir_celda(ws, r, 8, "", font=font_normal, alignment=align_left)
            aplicar_borde(ws, r, 1, 8)
            r += 1

        self._fila_insumos_fin = r - 1
        ws.freeze_panes = "A5"'''

with open('src/constructor.py', 'r', encoding='utf-8') as f:
    content = f.read()

old = '''def escribir_insumos(self, resultados_matcheo: list[dict], indice=None):
        """Crea la hoja Insumos: lista consolidada a nivel MATERIAL/insumo.

        Se agrega de todos los APUs (matrices): por cada insumo, la cantidad
        total de la obra = Σ (cantidad del insumo en el APU × cantidad de obra).

        Columnas:
        A: Clave | B: Tipo | C: Descripción | D: Unidad
        E: Costo unitario | F: Cantidad total | G: Importe | H: Notas
        """
        if "Insumos" in self.wb.sheetnames:
            del self.wb["Insumos"]
        ws = self.wb.create_sheet("Insumos")

        merge_y_escribir(ws, "A1:H1", "INSUMOS — Lista de materiales y recursos (explosión de APUs)",
                         font=font_titulo_blanco, fill=fill_azul_oscuro, alignment=align_center)
        for c in range(1, 9):
            ws.cell(row=1, column=c).fill = fill_azul_oscuro

        merge_y_escribir(ws, "A2:H2",
                         "Edita el Costo unitario (col E). Las hojas APU y el Presupuesto se recalculan.",
                         font=font_italica, alignment=align_left)

        headers = ["Clave", "Tipo", "Descripción", "Unidad",
                   "Costo unitario", "Cantidad total", "Importe", "Notas"]
        anchos = [14, 14, 55, 10, 16, 16, 18, 24]
        for i, (h, w) in enumerate(zip(headers, anchos), 1):
            escribir_celda(ws, 4, i, h, font=font_header_blanco, fill=fill_azul_medio,
                           alignment=align_center)
            ws.column_dimensions[col_letra(i)].width = w
        aplicar_borde(ws, 4, 1, 8)

        # Agregación por clave de insumo
        agg: dict[str, dict] = {}
        for res in resultados_matcheo:
            if res.get("estado") not in ("confirmado", "dudoso", "confirmado_ia"):
                continue
            clave = res.get("clave_sugerida")
            apu = self.apu_por_clave.get(clave)
            if not apu:
                continue
            cantidad_obra = float(res.get("cantidad", 0) or 0)
            for ins in apu.get("insumos", []):
                k = ins.get("clave")
                if not k:
                    continue
                a = agg.get(k)
                if a is None:
                    a = {"tipo": ins.get("tipo", ""), "desc": ins.get("descripcion", ""),
                         "unidad": ins.get("unidad", ""), "costo": float(ins.get("costo", 0) or 0),
                         "cant": 0.0}
                    agg[k] = a
                a["cant"] += float(ins.get("cantidad", 0) or 0) * cantidad_obra

        r = 5
        _lbl_tipo = {"material": "Material", "mano_obra": "Cuadrilla",
                      "equipo": "Equipo", "herramienta": "Herramienta",
                      "auxiliar": "Auxiliar", "maquinaria": "Maquinaria"}
        for k in sorted(agg.keys()):
            a = agg[k]
            escribir_celda(ws, r, 1, k, font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 2, _lbl_tipo.get(a["tipo"], a["tipo"]),
                           font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 3, a["desc"], font=font_normal, alignment=align_left)
            escribir_celda(ws, r, 4, a["unidad"], font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 5, round(a["costo"], 4), font=font_normal_rojo,
                           alignment=align_right, number_format="$#,##0.00")
            escribir_celda(ws, r, 6, round(a["cant"], 4), font=font_normal,
                           alignment=align_right, number_format="#,##0.0000")
            escribir_celda(ws, r, 7, f"=F{r}*E{r}", font=font_negrita,
                           alignment=align_right, number_format="$#,##0.00")
            escribir_celda(ws, r, 8, "", font=font_normal, alignment=align_left)
            aplicar_borde(ws, r, 1, 8)
            r += 1

        self._fila_insumos_fin = r - 1
        ws.freeze_panes = "A5"'''

new_code = '''def escribir_insumos(self, resultados_matcheo: list[dict], indice=None):
        """Crea la hoja Insumos: lista consolidada a nivel MATERIAL/insumo.

        Se agrega de todos los APUs (matrices): por cada insumo, la cantidad
        total de la obra = Σ (cantidad del insumo en el APU × cantidad de obra).

        Para tipo=cuadrilla: el Costo unitario es VLOOKUP a la hoja Cuadrillas (col C).
        Para otros tipos: costo directo editable.

        Columnas:
        A: Clave | B: Tipo | C: Descripción | D: Unidad
        E: Costo unitario | F: Cantidad total | G: Importe | H: Notas
        \"\"\"
        if "Insumos" in self.wb.sheetnames:
            del self.wb["Insumos"]
        ws = self.wb.create_sheet("Insumos")

        merge_y_escribir(ws, "A1:H1", "INSUMOS \u2014 Lista de materiales y recursos (explosi\u00f3n de APUs)",
                         font=DNA().font_titulo(), fill=DNA().fill_azul_oscuro(), alignment=DNA().align_center())
        for c in range(1, 9):
            ws.cell(row=1, column=c).fill = DNA().fill_azul_oscuro()

        merge_y_escribir(ws, "A2:H2",
                         "Edita el Costo unitario (col E). Para Cuadrillas, edita en hoja Cuadrillas (col C).",
                         font=DNA().font_italica(), alignment=DNA().align_left())

        headers = ["Clave", "Tipo", "Descripci\u00f3n", "Unidad",
                   "Costo unitario", "Cantidad total", "Importe", "Notas"]
        anchos = [14, 14, 55, 10, 16, 16, 18, 24]
        for i, (h, w) in enumerate(zip(headers, anchos), 1):
            escribir_celda(ws, 4, i, h, font=DNA().font_header(), fill=DNA().fill_header_azul(),
                           alignment=DNA().align_center())
            ws.column_dimensions[col_letra(i)].width = w
        aplicar_borde(ws, 4, 1, 8)

        # Agregación por clave de insumo
        agg: dict[str, dict] = {}
        for res in resultados_matcheo:
            if res.get("estado") not in ("confirmado", "dudoso", "confirmado_ia"):
                continue
            clave = res.get("clave_sugerida")
            apu = self.apu_por_clave.get(clave)
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
                         "cant": 0.0, "es_cuadrilla": es_cuadrilla}
                    agg[k] = a
                a["cant"] += float(ins.get("cantidad", 0) or 0) * cantidad_obra

        r = 5
        _lbl_tipo = {"material": "Material", "mano_obra": "Cuadrilla",
                      "equipo": "Equipo", "herramienta": "Herramienta",
                      "auxiliar": "Auxiliar", "maquinaria": "Maquinaria"}
        
        # Rango de la hoja Cuadrillas para VLOOKUP (col A=ID, col C=$/Jornal)
        rango_cuadrillas = "Cuadrillas!$A$4:$C$200"

        for k in sorted(agg.keys()):
            a = agg[k]
            es_cuadrilla = a.get("es_cuadrilla", False)
            
            escribir_celda(ws, r, 1, k, font=fuente_normal, alignment=align_center)
            escribir_celda(ws, r, 2, a.get("tipo", "material").capitalize(),
                           font=fuente_normal, alignment=align_center)
            escribir_celda(ws, r, 3, a["desc"], font=fuente_normal, alignment=align_left)
            escribir_celda(ws, r, 4, a["unidad"], font=fuente_normal, alignment=align_center)
            
            # Columna E: Costo unitario
            if a.get("es_cuadrilla", False):
                # Para cuadrillas: VLOOKUP a hoja Cuadrillas (col A=ID, col C=$/Jornal)
                formula_vlookup = f'=IFERROR(VLOOKUP(A{r},Cuadrillas!$A$4:$C$200,3,FALSE),0)'
                escribir_celda(ws, r, 5, formula_vlookup, font=fuente_normal_rojo,
                               alignment=align_right, number_format="$#,##0.00")
            else:
                # Materiales, equipos, etc.: costo directo editable
                escribir_celda(ws, r, 5, round(a["costo"], 4), font=fuente_normal_rojo,
                               alignment=align_right, number_format="$#,##0.00")
            
            escribir_celda(ws, r, 6, round(a["cant"], 4), font=fuente_normal,
                           alignment=align_right, number_format="#,##0.0000")
            escribir_celda(ws, r, 7, f"=F{r}*E{r}", font=fuente_negrita,
                           alignment=align_right, number_format="$#,##0.00")
            escribir_celda(ws, r, 8, "Cuadrilla" if a.get("es_cuadrilla") else "", 
                           font=fuente_normal, alignment=align_left)
            aplicar_borde(ws, r, 1, 8)
            r += 1

        self._fila_insumos_fin = r - 1
        ws.freeze_panes = "A5"'''

with open('src/constructor.py', 'r', encoding='utf-8') as f:
    content = f.read()

old = '''def escribir_insumos(self, resultados_matcheo: list[dict], indice=None):
        """Crea la hoja Insumos: lista consolidada a nivel MATERIAL/insumo.

        Se agrega de todos los APUs (matrices): por cada insumo, la cantidad
        total de la obra = Σ (cantidad del insumo en el APU × cantidad de obra).

        Columnas:
        A: Clave | B: Tipo | C: Descripción | D: Unidad
        E: Costo unitario | F: Cantidad total | G: Importe | H: Notas
        """
        if "Insumos" in self.wb.sheetnames:
            del self.wb["Insumos"]
        ws = self.wb.create_sheet("Insumos")

        merge_y_escribir(ws, "A1:H1", "INSUMOS — Lista de materiales y recursos (explosión de APUs)",
                         font=font_titulo_blanco, fill=fill_azul_oscuro, alignment=align_center)
        for c in range(1, 9):
            ws.cell(row=1, column=c).fill = fill_azul_oscuro

        merge_y_escribir(ws, "A2:H2",
                         "Edita el Costo unitario (col E). Las hojas APU y el Presupuesto se recalculan.",
                         font=font_italica, alignment=align_left)

        headers = ["Clave", "Tipo", "Descripción", "Unidad",
                   "Costo unitario", "Cantidad total", "Importe", "Notas"]
        anchos = [14, 14, 55, 10, 16, 16, 18, 24]
        for i, (h, w) in enumerate(zip(headers, anchos), 1):
            escribir_celda(ws, 4, i, h, font=font_header_blanco, fill=fill_azul_medio,
                           alignment=align_center)
            ws.column_dimensions[col_letra(i)].width = w
        aplicar_borde(ws, 4, 1, 8)

        # Agregación por clave de insumo
        agg: dict[str, dict] = {}
        for res in resultados_matcheo:
            if res.get("estado") not in ("confirmado", "dudoso", "confirmado_ia"):
                continue
            clave = res.get("clave_sugerida")
            apu = self.apu_por_clave.get(clave)
            if not apu:
                continue
            cantidad_obra = float(res.get("cantidad", 0) or 0)
            for ins in apu.get("insumos", []):
                k = ins.get("clave")
                if not k:
                    continue
                a = agg.get(k)
                if a is None:
                    a = {"tipo": ins.get("tipo", ""), "desc": ins.get("descripcion", ""),
                         "unidad": ins.get("unidad", ""), "costo": float(ins.get("costo", 0) or 0),
                         "cant": 0.0}
                    agg[k] = a
                a["cant"] += float(ins.get("cantidad", 0) or 0) * cantidad_obra

        r = 5
        _lbl_tipo = {"material": "Material", "mano_obra": "Cuadrilla",
                      "equipo": "Equipo", "herramienta": "Herramienta",
                      "auxiliar": "Auxiliar", "maquinaria": "Maquinaria"}
        for k in sorted(agg.keys()):
            a = agg[k]
            escribir_celda(ws, r, 1, k, font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 2, _lbl_tipo.get(a["tipo"], a["tipo"]),
                           font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 3, a["desc"], font=font_normal, alignment=align_left)
            escribir_celda(ws, r, 4, a["unidad"], font=font_normal, alignment=align_center)
            escribir_celda(ws, r, 5, round(a["costo"], 4), font=font_normal_rojo,
                           alignment=align_right, number_format="$#,##0.00")
            escribir_celda(ws, r, 6, round(a["cant"], 4), font=font_normal,
                           alignment=align_right, number_format="#,##0.0000")
            escribir_celda(ws, r, 7, f"=F{r}*E{r}", font=font_negrita,
                           alignment=align_right, number_format="$#,##0.00")
            escribir_celda(ws, r, 8, "", font=font_normal, alignment=align_left)
            aplicar_borde(ws, r, 1, 8)
            r += 1

        self._fila_insumos_fin = r - 1
        ws.freeze_panes = "A5"'''

new = '''def escribir_insumos(self, resultados_matcheo: list[dict], indice=None):
        """Crea la hoja Insumos: lista consolidada a nivel MATERIAL/insumo.

        Se agrega de todos los APUs (matrices): por cada insumo, la cantidad
        total de la obra = Σ (cantidad del insumo en el APU × cantidad de obra).

        Para tipo=cuadrilla: el Costo unitario es VLOOKUP a la hoja Cuadrillas (col C).
        Para otros tipos: costo directo editable.

        Columnas:
        A: Clave | B: Tipo | C: Descripción | D: Unidad
        E: Costo unitario | F: Cantidad total | G: Importe | H: Notas
        """
        if "Insumos" in self.wb.sheetnames:
            del self.wb["Insumos"]
        ws = self.wb.create_sheet("Insumos")

        merge_y_escribir(ws, "A1:H1", "INSUMOS — Lista de materiales y recursos (explosión de APUs)",
                         font=DNA().font_titulo(), fill=DNA().fill_azul_oscuro(), alignment=DNA().align_center())
        for c in range(1, 9):
            ws.cell(row=1, column=c).fill = DNA().fill_azul_oscuro()

        merge_y_escribir(ws, "A2:H2",
                         "Edita el Costo unitario (col E). Para Cuadrillas, edita en hoja Cuadrillas (col C).",
                         font=DNA().font_italica(), alignment=DNA().align_left())

        headers = ["Clave", "Tipo", "Descripción", "Unidad",
                   "Costo unitario", "Cantidad total", "Importe", "Notas"]
        anchos = [14, 14, 55, 10, 16, 16, 18, 24]
        for i, (h, w) in enumerate(zip(headers, anchos), 1):
            escribir_celda(ws, 4, i, h, font=DNA().font_header(), fill=DNA().fill_header_azul(),
                           alignment=DNA().align_center())
            ws.column_dimensions[col_letra(i)].width = w
        aplicar_borde(ws, 4, 1, 8)

        # Agregación por clave de insumo
        agg: dict[str, dict] = {}
        for res in resultados_matcheo:
            if res.get("estado") not in ("confirmado", "dudoso", "confirmado_ia"):
                continue
            clave = res.get("clave_sugerida")
            apu = self.apu_por_clave.get(clave)
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
                         "cant": 0.0, "es_cuadrilla": es_cuadrilla}
                    agg[k] = a
                a["cant"] += float(ins.get("cantidad", 0) or 0) * cantidad_obra

        r = 5
        _lbl_tipo = {"material": "Material", "mano_obra": "Cuadrilla",
                      "equipo": "Equipo", "herramienta": "Herramienta",
                      "auxiliar": "Auxiliar", "maquinaria": "Maquinaria"}
        
        # Rango de la hoja Cuadrillas para VLOOKUP (col A=ID, col C=$/Jornal)
        rango_cuadrillas = "Cuadrillas!$A$4:$C$200"

        for k in sorted(agg.keys()):
            a = agg[k]
            es_cuadrilla = a.get("es_cuadrilla", False)
            
            escribir_celda(ws, r, 1, k, font=fuente_normal, alignment=align_center)
            escribir_celda(ws, r, 2, a.get("tipo", "material").capitalize(),
                           font=fuente_normal, alignment=align_center)
            escribir_celda(ws, r, 3, a["desc"], font=fuente_normal, alignment=align_left)
            escribir_celda(ws, r, 4, a["unidad"], font=fuente_normal, alignment=align_center)
            
            # Columna E: Costo unitario
            if a.get("es_cuadrilla", False):
                # Para cuadrillas: VLOOKUP a hoja Cuadrillas (col A=ID, col C=$/Jornal)
                formula_vlookup = f'=IFERROR(VLOOKUP(A{r},Cuadrillas!$A$4:$C$200,3,FALSE),0)'
                escribir_celda(ws, r, 5, formula_vlookup, font=fuente_normal_rojo,
                               alignment=align_right, number_format="$#,##0.00")
            else:
                # Materiales, equipos, etc.: costo directo editable
                escribir_celda(ws, r, 5, round(a["costo"], 4), font=fuente_normal_rojo,
                               alignment=align_right, number_format="$#,##0.00")
            
            escribir_celda(ws, r, 6, round(a["cant"], 4), font=fuente_normal,
                           alignment=align_right, number_format="#,##0.0000")
            escribir_celda(ws, r, 7, f"=F{r}*E{r}", font=fuente_negrita,
                           alignment=align_right, number_format="$#,##0.00")
            escribir_celda(ws, r, 8, "Cuadrilla" if a.get("es_cuadrilla") else "", 
                           font=fuente_normal, alignment=align_left)
            aplicar_borde(ws, r, 1, 8)
            r += 1

        self._fila_insumos_fin = r - 1
        ws.freeze_panes = "A5"'''

if old in content:
    content = content.replace(old, new)
    with open('src/constructor.py', 'w', encoding='utf-8') as f:
        f.write(content)
    print("SUCCESS: escribir_insumos actualizado")
else:
    print("ERROR: No se encontró el código exacto para reemplazar")
    # Let's find what we have
    import re
    matches = list(re.finditer(r'def escribir_insumos', content))
    for m in matches:
        print(f"Found at position {m.start()}: {content[m.start():m.start()+100]}")