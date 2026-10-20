"""
Datos extraídos del TEXTO de boletines oficiales (PDF o HTML).

Para las series sin API estable (INEGI sin token, DANE, CNI, Fedesarrollo), cada
instituto publica un boletín mensual con una URL predecible, por ejemplo:
  INEGI  .../boletines/2026/enco/icc2026_10.pdf     (mes de publicación)
  DANE   .../operaciones/ISE/bol-ISE-jul2026.pdf    (mes de referencia)
El código descarga los últimos N boletines, busca la frase clave ("Indicador de
Confianza del Consumidor ... 45.1 puntos", "serie original ... 1,11 %") y arma la
serie con un dato por boletín. Así cada corrida completa los meses faltantes
aunque la API o el anexo Excel fallen. Los valores pasan por la misma validación
(rango + histórico) y se marcan en naranja ("verificar") salvo `verificado: true`.
"""
from __future__ import annotations

import datetime as dt
import io
import re
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urljoin

import pandas as pd

from .base import MONTHS, SourceResult, parse_number, register, strip_accents

MES3 = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
MES_LARGO = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
             "noviembre", "diciembre"]

_UNITS = {
    "%": r"(?:%|por ciento|por cento)",
    "puntos": r"(?:puntos|pontos|pts\.?|unidades)",
}
_NUMBER = r"(?<![\d.,])([-−–]?\s?\d{1,3}(?:[.,]\d{1,2})?)"
NEG_WORDS = (r"(ca[eí]|caen|cayeron|cay[oó]|retroced|disminu|baj[aóo]|bajaron|contra[ej]|contracci|negativ|"
             r"descend|recort|caida|reducci|queda|recu[oó]|ca[ií]ram|caiu|fell|drop|declin)")


def _num_regex(unidad: str) -> re.Pattern:
    if unidad in _UNITS:
        return re.compile(_NUMBER + r"\s*" + _UNITS[unidad], re.I)
    # "auto": con unidad, o número decimal suelto (exige decimales para no confundir días o años)
    return re.compile(_NUMBER + r"(?:\s*" + _UNITS["%"] + r"|\s*" + _UNITS["puntos"] + r"|(?=[\s,;.)]|$))", re.I)


def find_month(text: str, ref: Optional[dt.date] = None) -> Optional[Tuple[int, int]]:
    """Primer 'mes [de] año' del texto; si no hay año, se infiere con la fecha de referencia."""
    t = strip_accents(text).lower()
    rx = re.compile(r"\b(" + "|".join(MES_LARGO + ["setiembre", "janeiro", "fevereiro", "marco", "maio", "junho",
                                                    "julho", "setembro", "outubro", "novembro", "dezembro"]) +
                    r")\b(?:\s+(?:de|del)\s+|\s+|/|-)?(\d{4})?")
    m = rx.search(t)
    if not m:
        return None
    mo = MONTHS.get(m.group(1)) or (MES_LARGO.index(m.group(1)) + 1 if m.group(1) in MES_LARGO else None)
    if not mo:
        return None
    if m.group(2):
        return int(m.group(2)), mo
    ref = ref or dt.date.today()
    return (ref.year if mo <= ref.month else ref.year - 1), mo


def extract_value(text: str, patron: str, unidad: str = "auto", rango=(-1e9, 1e9), posicion: str = "despues",
                  ventana: int = 220) -> Optional[Tuple[float, int, str]]:
    """Busca `patron` y devuelve (valor, posición, ventana de texto) del número asociado."""
    plain = re.sub(r"\s+", " ", strip_accents(text.replace("\xa0", " ")))
    kw = re.compile(strip_accents(patron), re.I)
    numrx = _num_regex(unidad)
    lo, hi = rango
    for m in kw.finditer(plain):
        w0 = max(0, m.start() - 120)
        window = plain[w0: m.end() + ventana]
        kpos, kend = m.start() - w0, m.end() - w0
        cands = []
        for nm in numrx.finditer(window):
            raw = nm.group(1).replace("−", "-").replace("–", "-").replace(" ", "")
            if unidad == "auto" and not re.search(r"[.,]", raw) and not re.search(r"%|punt|pont", nm.group(0), re.I):
                continue
            val = parse_number(raw, "auto")
            if val is None:
                continue
            between = window[min(kend, nm.start()): nm.start()] if nm.start() > kend else window[max(0, nm.start() - 60): nm.start()]
            if not raw.startswith("-") and _negative_verb(window[max(0, nm.start() - 45): nm.start()]):
                val = -abs(val)
            if not (lo <= val <= hi):
                continue
            if posicion == "despues":
                if nm.start() < kend:
                    continue
                dist = nm.start() - kend
            else:
                dist = min(abs(nm.start() - kpos), abs(nm.end() - kpos))
            cands.append((dist, val, nm.start()))
        if cands:
            dist, val, pos = min(cands)
            return val, w0 + pos, window
    return None


def _negative_verb(prefix: str) -> bool:
    """'cayó 1.2%' es una variación negativa; 'bajó a 23,3' o 'se ubicó en' describen un NIVEL (sin signo)."""
    m = None
    for m in re.finditer(NEG_WORDS + r"\w*", prefix, re.I):
        pass
    if m is None:
        return False
    gap = prefix[m.end():]
    return not re.search(r"\b(a|al|en|hasta|de un nivel de|ubic|situ|registr)\b", gap, re.I)


def pdf_text(content: bytes, max_pages: int = 4) -> str:
    from pypdf import PdfReader  # import diferido: solo se necesita para boletines PDF
    reader = PdfReader(io.BytesIO(content))
    out = []
    for page in reader.pages[:max_pages]:
        try:
            out.append(page.extract_text() or "")
        except Exception:  # noqa: BLE001
            continue
    return "\n".join(out)


def html_text(content: bytes) -> str:
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(content, "lxml")
    for t in soup(["script", "style", "noscript"]):
        t.decompose()
    return soup.get_text(" ", strip=True)


def _months_back(n: int) -> List[dt.date]:
    today = dt.date.today()
    y, m = today.year, today.month
    out = []
    for _ in range(n):
        out.append(dt.date(y, m, 1))
        m -= 1
        if m == 0:
            y, m = y - 1, 12
    return out


def _fmt(tpl: str, d: dt.date) -> str:
    return tpl.format(anio=d.year, aa=str(d.year)[2:], mm=f"{d.month:02d}", m=d.month, mes=MES3[d.month - 1],
                      Mes=MES3[d.month - 1].capitalize(), mes_largo=MES_LARGO[d.month - 1])


def _shift(d: dt.date, months: int) -> dt.date:
    y, m = d.year, d.month - months
    while m < 1:
        y, m = y - 1, m + 12
    while m > 12:
        y, m = y + 1, m - 12
    return dt.date(y, m, 1)


def _documents(spec: Dict[str, Any], ctx, key: str):
    """Genera (url, fecha_supuesta_de_referencia) de los boletines a leer."""
    n = int(spec.get("meses_atras", 8))
    rezago = int(spec.get("rezago", 1))
    docs = []
    if spec.get("plantilla_url"):
        for d in _months_back(n + (rezago if spec.get("nombre_por", "publicacion") == "publicacion" else 0)):
            ref = _shift(d, rezago) if spec.get("nombre_por", "publicacion") == "publicacion" else d
            docs.append((_fmt(spec["plantilla_url"], d), ref))
    if spec.get("pagina"):
        from .web_files import _links
        try:
            page = spec["pagina"]
            html = ctx.http.text(f"{key}_pagina", page)
            for i, hop in enumerate(spec.get("saltos") or []):
                found = _links(html, page, hop)
                if not found:
                    break
                page = found[0]
                html = ctx.http.text(f"{key}_salto{i + 1}", page)
            for u in _links(html, page, spec.get("patron_enlace", r"\.pdf")):
                docs.append((u, None))
        except Exception as e:  # noqa: BLE001
            ctx.log.info("texto_web %s: página no disponible (%s)", key, e)
    if spec.get("url"):
        docs.append((spec["url"], None))
    return docs[: max(n, 1) * 2]


@register("texto_web")
def texto_web(spec: Dict[str, Any], ctx) -> SourceResult:
    key = spec.get("clave") or "texto_web"
    fmt = spec.get("formato", "pdf")
    found: Dict[pd.Timestamp, Tuple[float, str]] = {}
    errors = []
    for url, ref in _documents(spec, ctx, key):
        name = re.sub(r"\W+", "_", url.split("/")[-1].split("?")[0])[-60:]
        k = f"{key}_{name}"
        try:
            content = ctx.http.fetch(k, url, ext=fmt, reuse_cached=spec.get("cache_permanente", True))
        except Exception as e:  # noqa: BLE001
            errors.append(str(e)[:80])
            continue
        try:
            text = pdf_text(content, spec.get("paginas", 4)) if fmt == "pdf" else html_text(content)
        except Exception as e:  # noqa: BLE001
            errors.append(f"{url}: no se pudo leer ({e})")
            continue
        hit = extract_value(text, spec["patron"], spec.get("unidad", "auto"), tuple(spec.get("rango", (-1e9, 1e9))),
                            spec.get("posicion", "despues"))
        if not hit:
            continue
        val, pos, window = hit
        if ref is not None:
            # URL construida para un mes concreto: el boletín debe mencionar ese mes (si no, se descarta)
            head = strip_accents(text[:2500]).lower()
            if not re.search(rf"{MES_LARGO[ref.month - 1]}\W+(?:de\W+)?{ref.year}", head):
                errors.append(f"{url}: no menciona {MES_LARGO[ref.month - 1]} {ref.year}")
                continue
            ym = (ref.year, ref.month)
        else:
            ym = find_month(window) or find_month(text[:600])
        if ym is None:
            continue
        ts = pd.Timestamp(ym[0], ym[1], 1)
        found.setdefault(ts, (val, url))
    if not found:
        raise RuntimeError(f"Ningún boletín con '{spec['patron']}' ({'; '.join(errors[-2:])})")
    s = pd.Series({k: v[0] for k, v in found.items()}).sort_index()
    res = SourceResult(s, spec.get("etiqueta", "Boletín oficial (texto)"), list(found.values())[-1][1],
                       spec["patron"], verified=bool(spec.get("verificado", False)))
    res.notes.append(f"{len(s)} boletines leídos; dato extraído del texto")
    return res
