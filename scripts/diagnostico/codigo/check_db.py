import sqlite3, os
p = r"C:\Pontem_dev\costeo\costos.db"
print("existe:", os.path.exists(p))
c = sqlite3.connect(p)
cur = c.cursor()
tabs = [r[0] for r in cur.execute("SELECT name FROM sqlite_master WHERE type='table'")]
print("tablas:", tabs)
for t in ["conceptos", "apu_lineas", "apu_totales", "insumos"]:
    if t in tabs:
        n = cur.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
        print(f"  {t}: {n} filas")
# muestras
if "conceptos" in tabs:
    print("\nconcepto ejemplo:", cur.execute("SELECT id,clave,descripcion,unidad,precio_unitario,paquete FROM conceptos LIMIT 1").fetchone())
if "apu_lineas" in tabs and cur.execute("SELECT COUNT(*) FROM apu_lineas").fetchone()[0] > 0:
    print("apu_linea ejemplo:", cur.execute("SELECT * FROM apu_lineas LIMIT 1").fetchone())
    print("columnas apu_lineas:", [d[1] for d in cur.execute("PRAGMA table_info(apu_lineas)")])
else:
    print("\napu_lineas VACIO o no existe")
c.close()