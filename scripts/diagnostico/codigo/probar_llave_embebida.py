# -*- coding: utf-8 -*-
"""Prueba si la llave sk- embebida en el comentario de .secrets es valida.

NUNCA imprime la llave; solo reporta si la API la acepta.
"""
import sys
sys.stdout.reconfigure(encoding="utf-8")
import json
import re
from pathlib import Path

import requests

texto = (Path(".secrets") / "DEEPSEEK_API_KEY").read_text(encoding="utf-8", errors="ignore")
sk = re.findall(r"sk-[A-Za-z0-9]{20,}", texto)
if not sk:
    print("No hay token sk- embebido; nada que probar.")
    raise SystemExit(0)

url = "https://api.deepseek.com/v1/chat/completions"
headers = {"Authorization": f"Bearer {sk[0]}", "Content-Type": "application/json"}
payload = {
    "model": "deepseek-chat",
    "temperature": 0,
    "max_tokens": 30,
    "messages": [
        {"role": "system", "content": "Responde solo JSON."},
        {"role": "user", "content": '{"ping": true}'},
    ],
    "response_format": {"type": "json_object"},
}
try:
    r = requests.post(url, headers=headers, json=payload, timeout=60)
    print("HTTP status:", r.status_code)
    if r.status_code == 200:
        print("=> LA LLAVE EMBEBIDA ES VALIDA (probar de nuevo la guardaria funcionando)")
    else:
        print("cuerpo:", r.text[:300])
        print("=> la llave embebida NO es aceptada")
except Exception as e:
    print("error de red:", type(e).__name__, str(e)[:200])
