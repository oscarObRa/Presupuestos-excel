import openpyxl, zipfile, os
p = r"H:\Mi unidad\ICEL\PUs\Pretil OMA MTY\Pretiles_rev0.xlsx"

def cinfo(color):
    if color is None: return "none"
    if color.type == "rgb": return color.rgb
    if color.type == "theme": return f"theme{color.theme}@{color.tint:.2f}"
    return str(color.type)

def fill(c):
    f=c.fill
    if not f or f.patternType is None: return None
    return f"fill={cinfo(f.fgColor)}"

def font(c):
    fo=c.font
    return f"font(name={fo.name},sz={fo.sz},b={fo.bold},color={cinfo(fo.color)})"

wb = openpyxl.load_workbook(p)
print("HOJAS:", wb.sheetnames)
for name in ["PROPUESTA", "APU 01"]:
    if name not in wb.sheetnames: continue
    ws = wb[name]
    print("="*70); print("HOJA", name, ws.max_row, "x", ws.max_column, "| imgs:", len(ws._images), "| merged:", len(ws.merged_cells.ranges))
    for r in range(1, min(ws.max_row, 30)+1):
        vals=[]
        for c in range(1, min(ws.max_column, 12)+1):
            cell=ws.cell(row=r,column=c)
            if cell.value is not None or (fill(cell) or ""):
                extra = " ".join(x for x in [font(cell), fill(cell)] if x)
                vals.append(f"{cell.coordinate}={str(cell.value)[:22]!r} [{extra}]")
        if vals:
            print(f" F{r}: " + " | ".join(vals[:4]))
wb.close()
# imágenes
with zipfile.ZipFile(p) as z:
    print("MEDIA:", [n for n in z.namelist() if n.startswith("xl/media/")])