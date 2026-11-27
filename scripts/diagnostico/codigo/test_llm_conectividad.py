# -*- coding: utf-8 -*-
"""Prueba de conectividad con la IA sin tragar excepciones (llm.py las calla)."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, ".")
import json
import requests

from src.llm import cargar_config_llm, SYSTEM_PROMPT

cfg = cargar_config_llm("config.yaml")
url = f"{cfg['base_url'].rstrip('/')}/chat/completions"
headers = {
    "Authorization": f"Bearer {cfg['api_key']}",
    "Content-Type": "application/json",
}
payload = {
    "model": cfg["model"],
    "temperature": 0,
    "max_tokens": 60,
    "messages": [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": json.dumps(
            {"conceptos": [{"id": "TEST-1", "concepto": "prueba de conectividad",
                            "unidad": "pza", "cantidad": 1, "candidatos": []}]},
            ensure_ascii=False)},
    ],
    "response_format": {"type": "json_object"},
}
r = requests.post(url, headers=headers, json=payload, timeout=60)
print("HTTP status:", r.status_code)
print("cuerpo (primeros 600):", r.text[:600])
if r.status_code == 200:
    content = r.json()["choices"][0]["message"]["content"]
    print("content:", content[:300])
    print("=> CONECTIVIDAD REAL OK")
