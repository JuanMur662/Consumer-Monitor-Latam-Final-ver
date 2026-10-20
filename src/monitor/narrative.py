"""
Lectura automática por país.

1) Siempre: frases construidas con reglas a partir de los datos del Excel
   (nivel, cambio contra el mes previo y contra hace 12 meses) + el "por qué"
   extraído de los titulares.
2) Opcional: si hay ANTHROPIC_API_KEY en .env y resumen_llm != false, se pide a
   Claude un párrafo breve de síntesis por país usando SOLO esos datos y titulares.
"""
from __future__ import annotations

import json
import logging
import os
from typing import List, Optional

import requests

log = logging.getLogger(__name__)

MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]


def mes_txt(ts) -> str:
    return f"{MESES[ts.month - 1]}-{str(ts.year)[2:]}"


def fmt_val(v: Optional[float], unidad: str) -> str:
    if v is None:
        return "s/d"
    return f"{v * 100:.2f}%" if unidad == "pct" else f"{v:.1f}"


def fmt_delta(d: Optional[float], unidad: str) -> str:
    if d is None:
        return ""
    return f"{d * 100:+.2f} pp" if unidad == "pct" else f"{d:+.1f} pts"


def lectura(ind: dict) -> str:
    """Una frase por indicador."""
    u = ind["unidad"]
    v, p, a = ind.get("valor"), ind.get("previo"), ind.get("hace12")
    if v is None:
        return "sin dato reciente"
    txt = []
    if p is not None:
        d = v - p
        thr = 0.0005 if u == "pct" else 0.25
        verbo = "sube" if d > thr else ("baja" if d < -thr else "estable")
        txt.append(f"{verbo} ({fmt_delta(d, u)} vs mes previo)")
    if a is not None:
        d12 = v - a
        txt.append(f"{fmt_delta(d12, u)} vs hace 12 meses")
    if ind.get("umbral") is not None:
        txt.append("zona optimista (>50)" if v > ind["umbral"] else "zona pesimista (<50)")
    if ind.get("rezago_meses", 0) >= 3:
        txt.append(f"último dato con {ind['rezago_meses']} meses de rezago")
    return "; ".join(txt)


def rule_based(pais: str, indicadores: List[dict], noticias: list) -> str:
    partes = []
    for ind in indicadores:
        if ind.get("valor") is None:
            continue
        frase = f"{ind['nombre']} {fmt_val(ind['valor'], ind['unidad'])} en {ind['mes']}"
        if ind.get("previo") is not None:
            frase += f" ({fmt_delta(ind['valor'] - ind['previo'], ind['unidad'])} m/m)"
        partes.append(frase)
    txt = f"{pais}: " + "; ".join(partes) + "." if partes else f"{pais}: sin datos recientes."
    whys = [n for n in noticias if n.why][:2]
    if whys:
        txt += " Según la prensa: " + " / ".join(f"{n.tags[0] if n.tags else ''}: {n.why}" for n in whys) + "."
    return txt


def llm_summary(pais: str, indicadores: List[dict], noticias: list, model: str) -> Optional[str]:
    key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not key:
        return None
    datos = [{k: ind.get(k) for k in ("nombre", "mes", "valor", "previo", "hace12", "unidad")} for ind in indicadores]
    tit = [{"indicador": n.tags[0] if n.tags else "", "fecha": n.published.date().isoformat() if n.published else "",
            "fuente": n.source, "titular": n.title} for n in noticias[:15]]
    prompt = (
        "Eres analista de crédito corporativo (consumo/retail LatAm). Con SOLO estos datos y titulares, escribe en "
        f"español 2-3 frases para {pais}: qué se movió, por qué (según los titulares) y la implicación para el "
        "consumo/retail. No inventes cifras ni causas que no estén en los titulares. Valores 'pct' son fracciones "
        "(0.031 = 3.1%).\n\nDATOS:\n" + json.dumps(datos, ensure_ascii=False) +
        "\n\nTITULARES:\n" + json.dumps(tit, ensure_ascii=False)
    )
    try:
        r = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json"},
            json={"model": model, "max_tokens": 400, "messages": [{"role": "user", "content": prompt}]},
            timeout=60,
        )
        r.raise_for_status()
        return "".join(b.get("text", "") for b in r.json().get("content", [])).strip()
    except Exception as e:  # noqa: BLE001
        log.warning("Resumen LLM no disponible para %s: %s", pais, e)
        return None
