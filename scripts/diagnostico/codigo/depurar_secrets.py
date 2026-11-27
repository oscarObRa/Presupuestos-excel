# -*- coding: utf-8 -*-
"""Verifica las llaves de .secrets SIN imprimirlas (solo diagnostico de forma)."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
from pathlib import Path

for nombre in ("DEEPSEEK_API_KEY", "OPENROUTER_API_KEY"):
    p = Path(".secrets") / nombre
    if not p.exists():
        print(f"{nombre}: NO EXISTE")
        continue
    texto = p.read_text(encoding="utf-8", errors="ignore")
    lineas = texto.splitlines()
    n_comentarios = sum(1 for l in lineas if l.strip().startswith("#"))
    no_comentario = [l.strip() for l in lineas if l.strip() and not l.strip().startswith("#")]
    # candidato sk- embebido en comentarios
    import re
    sk = re.findall(r"sk-[A-Za-z0-9]{20,}", texto)
    print(f"{nombre}: {len(lineas)} lineas, {n_comentarios} comentarios, "
          f"{len(no_comentario)} lineas de valor, contiene token sk-: {bool(sk)}")
    if sk:
        print(f"   token candidato: sk-...{sk[0][-6:]} (len {len(sk[0])})")
    if no_comentario:
        print(f"   primera linea de valor: {no_comentario[0][:12]}...{no_comentario[0][-4:]} (len {len(no_comentario[0])})")
