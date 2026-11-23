import sys
sys.path.insert(0, ".")
from src.apu_source import parsear_archivo

archivos = [
    r"G:\Mi unidad\Costos\PU2026\OBRA CIVIL, EDIFICACION 2026\MATRICES\PRELIMINARES\Excavaciones a mano.xlsx",
]
import glob
# añadir un par de archivos de otras especialidades
for pat in [
    r"G:\Mi unidad\Costos\PU2026\OBRA CIVIL, EDIFICACION 2026\MATRICES\ALBAÑILERIA\*.xlsx",
    r"G:\Mi unidad\Costos\PU2026\OBRA CIVIL, EDIFICACION 2026\MATRICES\CIMENTACIONES\*.xlsx",
]:
    fs = glob.glob(pat)
    if fs:
        archivos.append(fs[0])

for p in archivos:
    apus = parsear_archivo(p)
    print("=" * 70)
    print(p.split("\\")[-1], "->", len(apus), "APUs")
    for a in apus[:2]:
        print(f"  {a['clave']} | {a['unidad']} | PU={a['pu']} | {a['descripcion'][:45]}")
        tipos = {}
        for ins in a["insumos"]:
            tipos[ins["tipo"]] = tipos.get(ins["tipo"], 0) + 1
        print("     insumos por tipo:", tipos)
        for ins in a["insumos"][:4]:
            print("       ", ins["tipo"], "|", ins["clave"], "|", ins["descripcion"][:28], "|", ins["cantidad"], "x", ins["costo"])