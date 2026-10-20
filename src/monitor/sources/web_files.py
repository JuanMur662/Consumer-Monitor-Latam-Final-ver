"""
Fuentes que publican la serie como archivo Excel enlazado desde una página web
(DANE, CNI, INE Chile, Fedesarrollo). El nombre del archivo cambia cada mes, así que:
  1. se abre la página y se busca el enlace con `patron_enlace` (se pueden encadenar "saltos")
  2. si no aparece, se prueban URLs construidas con `plantilla_url` para los últimos meses
  3. se lee el archivo con el lector genérico (tables.py) buscando `fila_regex`
"""
from __future__ import annotations

import datetime as dt
import re
from typing import Any, Dict, List, Optional
from urllib.parse import urljoin

from .base import SourceResult, register
from .tables import describe, find_series, read_all_sheets

MES3 = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
MES_PT = ["janeiro", "fevereiro", "marco", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro",
          "novembro", "dezembro"]


def _links(html: str, base: str, pattern: str) -> List[str]:
    rx = re.compile(pattern, re.I)
    out = []
    for href in re.findall(r"""href\s*=\s*["']([^"']+)["']""", html, re.I):
        if rx.search(href):
            u = urljoin(base, href.replace("&amp;", "&"))
            if u not in out:
                out.append(u)
    # enlaces en texto plano / JSON embebido
    for m in re.finditer(r"""https?://[^\s"'<>]+""", html):
        if rx.search(m.group(0)) and m.group(0) not in out:
            out.append(m.group(0))
    return out


def _template_urls(tpl: str, months_back: int = 5) -> List[str]:
    today = dt.date.today()
    urls = []
    y, m = today.year, today.month
    for _ in range(months_back):
        urls.append(tpl.format(mes=MES3[m - 1], Mes=MES3[m - 1].capitalize(), mm=f"{m:02d}", anio=y,
                               aa=str(y)[2:], mes_pt=MES_PT[m - 1]))
        m -= 1
        if m == 0:
            y, m = y - 1, 12
    return urls


def discover_file(spec: Dict[str, Any], ctx, key: str) -> (bytes, str):
    errors = []
    candidates: List[str] = []
    if spec.get("pagina"):
        try:
            page = spec["pagina"]
            html = ctx.http.text(f"{key}_pagina", page)
            hops = spec.get("saltos") or []
            for i, hop in enumerate(hops):
                found = _links(html, page, hop)
                if not found:
                    raise RuntimeError(f"salto {i + 1} sin enlaces que cumplan {hop}")
                page = found[0]
                html = ctx.http.text(f"{key}_salto{i + 1}", page)
            candidates += _links(html, page, spec["patron_enlace"])
        except Exception as e:  # noqa: BLE001
            errors.append(f"página: {e}")
    if spec.get("plantilla_url"):
        candidates += _template_urls(spec["plantilla_url"], spec.get("meses_atras", 5))
    if spec.get("url"):
        candidates.append(spec["url"])
    seen = set()
    for url in candidates:
        if url in seen:
            continue
        seen.add(url)
        ext = "xls" if url.lower().split("?")[0].endswith(".xls") else "xlsx"
        try:
            content = ctx.http.fetch(f"{key}_archivo", url, ext=ext)
            if content[:2] == b"PK" or content[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
                return content, url
            errors.append(f"{url}: no es un Excel")
        except Exception as e:  # noqa: BLE001
            errors.append(f"{url}: {e}")
    raise RuntimeError("No se pudo obtener el archivo. " + " | ".join(errors[-4:]))


@register("archivo_web")
def archivo_web(spec: Dict[str, Any], ctx) -> SourceResult:
    key = spec.get("clave") or re.sub(r"\W+", "_", spec.get("patron_enlace", "archivo"))[:50]
    content, url = discover_file(spec, ctx, key)
    ext = "xls" if url.lower().split("?")[0].endswith(".xls") else "xlsx"
    sheets = read_all_sheets(content, ext)
    cands = find_series(sheets, spec["fila_regex"], spec.get("hoja_regex"), spec.get("excluir_regex"))
    if not cands:
        if ctx.diagnostico:
            ctx.log.info("Estructura del archivo %s:\n%s", url, describe(sheets))
        raise RuntimeError(f"No encontré '{spec['fila_regex']}' en {url} (corre `diagnostico` para ver la estructura)")
    best = cands[0]
    res = SourceResult(best.series, spec.get("etiqueta", "Archivo web"), url, f"{best.sheet} / {best.label} ({best.layout})")
    res.notes.append(f"{len(cands)} coincidencias; se usó hoja '{best.sheet}' disposición {best.layout}")
    # el pipeline valida también estas filas y se queda con la que coincide con el histórico del Excel
    res.alternatives = [(f"{c.sheet} / {c.label[:60]} ({c.layout})", c.series) for c in cands[1:25]]
    return res
