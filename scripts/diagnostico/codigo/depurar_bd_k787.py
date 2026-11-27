# -*- coding: utf-8 -*-
"""Verifica en la BD canónica el APU de K78.7 y los precios de sus insumos."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
import sqlite3

DB = r"C:/Pontem_dev/costeo/dist/contrato_costos/1.2.0/costos_canonico.db"
conn = sqlite3.connect("file:" + DB + "?mode=ro", uri=True)
conn.row_factory = sqlite3.Row

print("=== concepto K78.7 (PU oficial) ===")
for f in conn.execute("""SELECT clave, descripcion, nivel_2, unidad,
                         precio_unitario_manual, precio_unitario_publicado
                         FROM conceptos WHERE clave='K78.7'"""):
    print({k: f[k] for k in f.keys()})

print()
print("=== apu_lineas del APU de K78.7 (via v_apu_linea_costo / apu_lineas) ===")
try:
    for f in conn.execute("""SELECT * FROM v_apu_linea_costo
                             WHERE apu_id IN (SELECT apu_id FROM v_concepto_pu WHERE clave='K78.7')
                             LIMIT 12"""):
        print({k: f[k] for k in f.keys()})
except Exception as e:
    print("v_apu_linea_costo:", e)
    for f in conn.execute("""SELECT al.* FROM apu_lineas al
                             JOIN apus a ON al.apu_id = a.id
                             JOIN conceptos c ON c.apu_id = a.id
                             WHERE c.clave='K78.7'"""):
        print({k: f[k] for k in f.keys()})

print()
print("=== Precios de los insumos de ese APU ===")
for f in conn.execute("""SELECT i.clave, i.descripcion, i.unidad, ip.precio, ip.fuente
                         FROM apu_lineas al
                         JOIN apus a ON al.apu_id = a.id
                         JOIN conceptos c ON c.apu_id = a.id
                         JOIN insumos i ON i.id = al.insumo_id
                         LEFT JOIN insumo_precios ip ON ip.insumo_id = i.id
                         WHERE c.clave='K78.7'"""):
    print({k: f[k] for k in f.keys()})
conn.close()
