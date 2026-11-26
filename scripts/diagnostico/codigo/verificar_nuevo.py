import openpyxl
p = r"H:\Mi unidad\ICEL\PUs\CASA VENADO\presupuesto.xlsx"
wb = openpyxl.load_workbook(p)
names = wb.sheetnames
print("Total hojas:", len(names))
print("Fijas:", [n for n in names if not n.startswith("APU_")][:15])
apus = [n for n in names if n.startswith("APU_")]
print("Hojas APU:", len(apus), "ej:", apus[:5])
print("¿Revisión?", "Revision" in names, "| ¿Explosion?", "Explosion" in names)

ws = wb["Insumos"]
print("\n=== INSUMOS (primeras 12 filas) ===")
for r in range(4, 16):
    vals = [ws.cell(row=r, column=c).value for c in range(1, 8)]
    print("  ", vals)

if apus:
    wa = wb[apus[0]]
    print(f"\n=== {apus[0]} (primeras 18 filas) ===")
    for r in range(1, 19):
        vals = [wa.cell(row=r, column=c).value for c in range(1, 7)]
        if any(v not in (None, "") for v in vals):
            print("  ", vals)
wb.close()