"""
llm.py
======
Resuelve conceptos dudosos/sin match usando un LLM (API OpenAI-compatible).

Proveedores soportados (ambos OpenAI-compatible):
- DeepSeek directo:   base_url https://api.deepseek.com/v1   (DEEPSEEK_API_KEY)
- OpenRouter:         base_url https://openrouter.ai/api/v1  (OPENROUTER_API_KEY)

NO es un chat: el script arma un prompt con la descripción del concepto y sus
mejores candidatos del catálogo, y el modelo devuelve JSON con la clave elegida.
Las sugerencias se escriben en la hoja `Revisión` (y, si `auto_aplicar`, se
usan para el presupuesto).

Uso:
    from src.llm import ClienteLLM, cargar_config_llm, aplicar_sugerencias

    cfg = cargar_config_llm("config.yaml")
    llm = ClienteLLM(cfg)
    if llm.disponible:
        sugerencias = llm.resolver(resultados)          # list[dict]
        aplicar_sugerencias(resultados, indice, sugerencias, cfg.get("auto_aplicar", True))
"""

from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Optional

try:
    import requests
except ImportError:  # pragma: no cover
    requests = None


# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

PROVEEDORES = {
    "deepseek": {
        "base_url": "https://api.deepseek.com/v1",
        "model": "deepseek-chat",
        "key_env": "DEEPSEEK_API_KEY",
    },
    "openrouter": {
        "base_url": "https://openrouter.ai/api/v1",
        "model": "deepseek/deepseek-chat",
        "key_env": "OPENROUTER_API_KEY",
    },
}

# Modelos sugeridos (para mostrar / documentar)
MODELOS_RECOMENDADOS = {
    "deepseek": ["deepseek-chat", "deepseek-reasoner"],
    "openrouter": ["deepseek/deepseek-chat", "deepseek/deepseek-r1"],
}


def _leer_env_archivo(ruta: str = ".env") -> dict:
    """Lee un .env simple (KEY=VALUE) sin dependencias.
    También lee archivos individuales en .secrets/ (un archivo por variable)."""
    datos = {}
    # 1. Archivo .env tradicional
    p = Path(ruta)
    if p.exists():
        for linea in p.read_text(encoding="utf-8", errors="ignore").splitlines():
            linea = linea.strip()
            if not linea or linea.startswith("#") or "=" not in linea:
                continue
            k, v = linea.split("=", 1)
            datos[k.strip()] = v.strip().strip('"').strip("'")
    
    # 2. Archivos individuales en .secrets/ (un archivo por variable)
    secrets_dir = Path(".secrets")
    if secrets_dir.exists() and secrets_dir.is_dir():
        for archivo in secrets_dir.iterdir():
            if archivo.is_file() and not archivo.name.startswith("."):
                try:
                    valor = archivo.read_text(encoding="utf-8", errors="ignore").strip()
                    # El nombre del archivo es la variable (ej. DEEPSEEK_API_KEY)
                    datos[archivo.name] = valor.strip().strip('"').strip("'")
                except Exception:
                    pass
    
    return datos


def cargar_config_llm(config_path: str = "config.yaml") -> dict:
    """Carga la sección `llm` de config.yaml, con overrides de entorno/.env."""
    cfg = {}
    try:
        import yaml
        p = Path(config_path)
        if p.exists():
            data = yaml.safe_load(p.read_text(encoding="utf-8")) or {}
            cfg = dict(data.get("llm", {}) or {})
    except Exception:
        cfg = {}

    env = _leer_env_archivo(".env")

    provider = (os.environ.get("LLM_PROVIDER") or env.get("LLM_PROVIDER")
                or cfg.get("provider") or "deepseek").strip().lower()
    if provider not in PROVEEDORES:
        provider = "deepseek"
    prov = PROVEEDORES[provider]

    model = (os.environ.get("LLM_MODEL") or env.get("LLM_MODEL")
             or cfg.get("model") or prov["model"])
    base_url = (os.environ.get("LLM_BASE_URL") or env.get("LLM_BASE_URL")
                or cfg.get("base_url") or prov["base_url"])
    key_env = (cfg.get("api_key_env") or prov["key_env"])

    api_key = (os.environ.get("LLM_API_KEY") or env.get("LLM_API_KEY")
               or os.environ.get(key_env) or env.get(key_env) or "")

    return {
        "provider": provider,
        "model": str(model),
        "base_url": str(base_url),
        "api_key": api_key,
        "api_key_env": key_env,
        "max_tokens": int(cfg.get("max_tokens", 2048)),
        "temperature": float(cfg.get("temperature", 0)),
        "lote": int(cfg.get("lote", 8)),
        "max_items": int(cfg.get("max_items", 200)),
        "max_candidatos": int(cfg.get("max_candidatos", 8)),
        "auto_aplicar": bool(cfg.get("auto_aplicar", True)),
    }


# ---------------------------------------------------------------------------
# Prompt
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = (
    "Eres un ingeniero experto en presupuestos de obra en México. "
    "Recibes conceptos de un presupuesto (con su descripción, unidad y cantidad) "
    "y una lista de candidatos de un catálogo de precios unitarios. "
    "Tu tarea: elegir la clave del catálogo que MEJOR corresponde a cada concepto. "
    "Reglas estrictas:\n"
    "1) SOLO puedes usar claves de la lista de candidatos. Nunca inventes claves.\n"
    "2) Si ningún candidato corresponde realmente, usa clave null.\n"
    "3) Respeta la unidad: no asignes un concepto en pza a uno medido en ml/kg/m2/m3.\n"
    "4) Prefiere el candidato más específico; evita conceptos que engloban un sistema completo.\n"
    "5) Responde ÚNICAMENTE con JSON válido, sin texto adicional.\n"
    'Formato exacto: {"resultados":[{"id":"<id>","clave":<clave|null>,'
    '"confianza":<0.0-1.0>,"motivo":"<breve>"}]}'
)


def _construir_user_prompt(items: list[dict]) -> str:
    """items: [{id, concepto, unidad, cantidad, candidatos:[...]}]"""
    return json.dumps({"conceptos": items}, ensure_ascii=False, indent=None)


# ---------------------------------------------------------------------------
# Cliente
# ---------------------------------------------------------------------------

class ClienteLLM:
    def __init__(self, cfg: dict):
        self.cfg = cfg
        self.provider = cfg.get("provider", "deepseek")
        self.model = cfg.get("model", "")
        self.base_url = (cfg.get("base_url") or "").rstrip("/")
        self.api_key = cfg.get("api_key", "")

    @property
    def disponible(self) -> bool:
        return bool(requests is not None and self.api_key and self.base_url and self.model)

    def estado(self) -> str:
        if requests is None:
            return "IA: no disponible (falta 'requests')"
        if not self.api_key:
            return f"IA: no disponible (falta {self.cfg.get('api_key_env','API key')})"
        return f"IA: {self.provider} / {self.model}"

    # -- llamada --------------------------------------------------------
    def _llamar(self, items: list[dict]) -> list[dict]:
        url = f"{self.base_url}/chat/completions"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        # OpenRouter pide headers opcionales de referencia (no obligatorios)
        if self.provider == "openrouter":
            headers["HTTP-Referer"] = "https://pontem.dev"
            headers["X-Title"] = "Presupuestador PONTEM"

        payload = {
            "model": self.model,
            "temperature": self.cfg.get("temperature", 0),
            "max_tokens": self.cfg.get("max_tokens", 2048),
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": _construir_user_prompt(items)},
            ],
            "response_format": {"type": "json_object"},
        }
        try:
            r = requests.post(url, headers=headers, json=payload, timeout=90)
        except Exception:
            return []

        if r.status_code != 200:
            return []

        try:
            data = r.json()
            content = data["choices"][0]["message"]["content"]
            content = _limpiar_json(content)
            parsed = json.loads(content)
            res = parsed.get("resultados", parsed if isinstance(parsed, list) else [])
            return res if isinstance(res, list) else []
        except Exception:
            return []

    # -- resolución -----------------------------------------------------
    def resolver(self, resultados: list[dict]) -> list[dict]:
        """Resuelve los conceptos dudosos/sin match. Devuelve sugerencias."""
        if not self.disponible:
            return []

        # candidatos por id de entrada
        pendientes = []
        for r in resultados:
            if r.get("estado") not in ("dudoso", "sin_match"):
                continue
            cands = r.get("candidatos") or []
            if not cands:
                continue
            pendientes.append(r)
            if len(pendientes) >= self.cfg.get("max_items", 200):
                break

        if not pendientes:
            return []

        max_cand = self.cfg.get("max_candidatos", 8)
        lote = max(1, self.cfg.get("lote", 8))
        sugerencias: list[dict] = []

        for i in range(0, len(pendientes), lote):
            chunk = pendientes[i:i + lote]
            items = []
            for r in chunk:
                items.append({
                    "id": r["id_entrada"],
                    "concepto": r["concepto_entrada"],
                    "unidad": r["unidad_entrada"],
                    "cantidad": r["cantidad"],
                    "candidatos": [
                        {
                            "clave": c.get("clave"),
                            "descripcion": c.get("descripcion"),
                            "unidad": c.get("unidad"),
                            "precio_unitario": c.get("precio_unitario"),
                        }
                        for c in (r.get("candidatos") or [])[:max_cand]
                    ],
                })
            resp = self._llamar(items)
            sugerencias.extend(resp)

        return sugerencias


# ---------------------------------------------------------------------------
# Parseo y aplicación
# ---------------------------------------------------------------------------

def _limpiar_json(texto: str) -> str:
    """Quita cercas ```json ... ``` si el modelo las incluye."""
    if texto is None:
        return "{}"
    t = texto.strip()
    t = re.sub(r"^```(?:json)?", "", t, flags=re.IGNORECASE).strip()
    t = re.sub(r"```$", "", t).strip()
    return t


def aplicar_sugerencias(resultados: list[dict], indice, sugerencias: list[dict],
                        auto_aplicar: bool = True) -> int:
    """Escribe las sugerencias de la IA en los resultados.

    - Guarda clave_ia, confianza_ia, motivo_ia (para la hoja Revisión).
    - Si auto_aplicar y la clave es válida, actualiza clave_sugerida / PU / estado.

    Devuelve cuántas sugerencias aplicadas al presupuesto.
    """
    if not sugerencias:
        return 0

    por_id = {}
    for s in sugerencias:
        if isinstance(s, dict) and s.get("id") is not None:
            por_id[str(s.get("id"))] = s

    aplicadas = 0
    for r in resultados:
        s = por_id.get(str(r.get("id_entrada")))
        if not s:
            continue
        clave = s.get("clave")
        conf = s.get("confianza", 0)
        motivo = s.get("motivo", "")

        # Validar que la clave exista y esté entre los candidatos del resultado
        claves_cand = {c.get("clave") for c in (r.get("candidatos") or [])}
        valida = bool(clave) and clave in claves_cand and indice.por_clave(clave) is not None

        r["clave_ia"] = clave if valida else None
        r["confianza_ia"] = conf
        r["motivo_ia"] = motivo

        if auto_aplicar and valida:
            c = indice.por_clave(clave)
            r["clave_sugerida"] = clave
            r["descripcion_sugerida"] = c.get("descripcion", "")
            r["unidad_sugerida"] = c.get("unidad", "")
            r["pu_sugerido"] = c.get("precio_unitario", 0)
            r["estado"] = "confirmado_ia"
            r["metodo"] = "llm"
            aplicadas += 1

    return aplicadas