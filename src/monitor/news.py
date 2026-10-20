"""
Noticias: titulares recientes por país/indicador (Google News RSS y feeds RSS
oficiales configurables) + extracción simple de "por qué" se movió el dato.

No requiere API keys. Si existe ANTHROPIC_API_KEY en .env, narrative.py puede
redactar además un párrafo de síntesis por país (opcional).
"""
from __future__ import annotations

import datetime as dt
import email.utils
import html
import re
from dataclasses import dataclass, field
from typing import List, Optional
from urllib.parse import quote_plus

from lxml import etree

from .sources.base import MONTHS, SourceResult, parse_number, register, strip_accents, to_monthly

LANGS = {
    "es-419": ("es-419", "es-419"),
    "pt-BR": ("pt-BR", "pt-419"),
    "en-US": ("en-US", "en"),
}


@dataclass
class NewsItem:
    title: str
    link: str
    source: str
    published: Optional[dt.datetime]
    summary: str = ""
    query: str = ""
    why: str = ""
    tags: list = field(default_factory=list)


def _clean(txt: str) -> str:
    txt = html.unescape(re.sub(r"<[^>]+>", " ", txt or ""))
    return re.sub(r"\s+", " ", txt).strip()


def parse_rss(content: bytes, query: str = "") -> List[NewsItem]:
    root = etree.fromstring(content, parser=etree.XMLParser(recover=True))
    items = []
    for it in root.iter("item"):
        title = _clean(it.findtext("title") or "")
        src_el = it.find("source")
        source = _clean(src_el.text) if src_el is not None and src_el.text else ""
        if source and title.endswith(" - " + source):
            title = title[: -(len(source) + 3)]
        pub = it.findtext("pubDate")
        try:
            published = email.utils.parsedate_to_datetime(pub).replace(tzinfo=None) if pub else None
        except Exception:  # noqa: BLE001
            published = None
        items.append(NewsItem(title, (it.findtext("link") or "").strip(), source, published,
                              _clean(it.findtext("description") or ""), query))
    # Atom (algunos feeds oficiales)
    ns = {"a": "http://www.w3.org/2005/Atom"}
    for e in root.findall(".//a:entry", ns):
        link = e.find("a:link", ns)
        pub = e.findtext("a:updated", namespaces=ns) or e.findtext("a:published", namespaces=ns)
        try:
            published = dt.datetime.fromisoformat(pub[:19]) if pub else None
        except Exception:  # noqa: BLE001
            published = None
        items.append(NewsItem(_clean(e.findtext("a:title", namespaces=ns) or ""),
                              link.get("href") if link is not None else "", "", published,
                              _clean(e.findtext("a:summary", namespaces=ns) or ""), query))
    return items


def google_news(ctx, query: str, lang: str = "es-419", country: str = "CO", days: int = 45) -> List[NewsItem]:
    hl, ceid_lang = LANGS.get(lang, (lang, lang))
    q = quote_plus(f"{query} when:{days}d")
    url = f"https://news.google.com/rss/search?q={q}&hl={hl}&gl={country}&ceid={country}:{ceid_lang}"
    key = "gnews_" + re.sub(r"\W+", "_", strip_accents(f"{country}_{query}"))[:80]
    content = ctx.http.fetch(key, url, ext="xml")
    return parse_rss(content, query)


def rss_feed(ctx, url: str, name: str) -> List[NewsItem]:
    content = ctx.http.fetch("rss_" + re.sub(r"\W+", "_", name)[:60], url, ext="xml")
    items = parse_rss(content, name)
    for it in items:
        it.source = it.source or name
    return items


# ---------------------------------------------------------------- "por qué"
CAUSAL = [
    r"impulsad[oa]s? por", r"jalonad[oa]s? por", r"explicad[oa]s? por", r"debido a", r"gracias a",
    r"ante", r"tras", r"por (?:el|la|los|las|un|una|mayores|menores|menor|mayor)", r"presionad[oa]s? por",
    r"lastrad[oa]s? por", r"afectad[oa]s? por", r"liderad[oa]s? por", r"en medio de", r"a causa de",
    r"puxad[oa]s? por", r"impulsionad[oa]s? por", r"devido a", r"com alta de", r"com queda de",
    r"driven by", r"due to", r"amid", r"led by", r"on the back of",
]
_CAUSAL_RX = re.compile(r"\b(" + "|".join(CAUSAL) + r")\b\s+([^.;:|]{6,140})", re.I)


def extract_why(text: str) -> str:
    m = _CAUSAL_RX.search(text or "")
    if not m:
        return ""
    clause = m.group(2).strip(" ,-–")
    clause = re.split(r"\s[-–|]\s", clause)[0]
    return f"{m.group(1)} {clause}"


def rank_items(items: List[NewsItem], must: Optional[str], exclude: Optional[str], max_items: int,
               max_age_days: int = 60) -> List[NewsItem]:
    rx = re.compile(strip_accents(must), re.I) if must else None
    xrx = re.compile(strip_accents(exclude), re.I) if exclude else None
    now = dt.datetime.now()
    seen, out = set(), []
    for it in sorted(items, key=lambda x: x.published or dt.datetime(1900, 1, 1), reverse=True):
        text = strip_accents(it.title + " " + it.summary)
        if rx and not rx.search(text):
            continue
        if xrx and xrx.search(text):
            continue
        if it.published and (now - it.published).days > max_age_days:
            continue
        k = re.sub(r"\W+", "", strip_accents(it.title).lower())[:60]
        if k in seen:
            continue
        seen.add(k)
        it.why = extract_why(it.title) or extract_why(it.summary)
        out.append(it)
        if len(out) >= max_items:
            break
    return out


# ---------------------------------------------------------------- dato desde titulares
NEG_WORDS = r"(ca[eí]|caen|cayeron|cay[oó]|retroced|disminu|baj[aóo]|bajaron|contra[ej]|contracci|negativ|descend|recort|caída|caida|queda|recua|cai|ca[ií]ram|fell|drop|declin)"
_NUM = r"(-?\d{1,3}(?:[.,]\d{1,2})?)\s*(?:%|por ciento|pontos|puntos|pts)"


def _infer_month(text: str, published: Optional[dt.datetime]):
    t = strip_accents(text).lower()
    best = None
    for name, num in MONTHS.items():
        if len(name) < 4 or name.endswith("_pt"):
            continue
        for m in re.finditer(r"\b" + name + r"\b(?:\s+(?:de|del)\s+)?\s*(\d{4})?", t):
            year = int(m.group(1)) if m.group(1) else None
            if best is None or m.start() < best[2]:
                best = (num, year, m.start())
    pub = published or dt.datetime.now()
    if best is None:
        prev = (pub.replace(day=1) - dt.timedelta(days=1))
        return prev.year, prev.month, False
    num, year, _ = best
    if year is None:
        year = pub.year if num <= pub.month else pub.year - 1
    return year, num, True


@register("noticia_dato")
def noticia_dato(spec, ctx) -> SourceResult:
    """Extrae el último dato publicado desde titulares (respaldo para fuentes sin API: ANTAD, Fedesarrollo)."""
    items = google_news(ctx, spec["consulta"], spec.get("idioma", "es-419"), spec.get("pais", "MX"),
                        spec.get("dias", 40))
    from .sources.text_values import extract_value
    lo, hi = spec.get("rango", [-1e9, 1e9])
    found = {}
    for it in items:
        text = it.title + ". " + it.summary
        # el número más cercano a la palabra clave (evita tomar "tiendas totales" en vez de "tiendas iguales");
        # unidad "auto" acepta "bajó a 23,3" (sin %) además de "2.3%" o "45.1 puntos"
        hit = extract_value(text, spec["patron"], spec.get("unidad", "auto"), (lo, hi),
                            spec.get("posicion", "cercano"), ventana=160)
        if not hit:
            continue
        val, pos, window = hit
        y, mo, explicit = _infer_month(window, it.published)
        key = (y, mo, round(val, 1))
        found.setdefault(key, set()).add(it.source or it.link)
    if not found:
        raise RuntimeError(f"Sin titulares con dato para '{spec['consulta']}'")
    latest = max((k[0], k[1]) for k in found)
    groups = {k: v for k, v in found.items() if (k[0], k[1]) == latest}
    best_key = max(groups, key=lambda k: len(groups[k]))
    n_src = len(groups[best_key])
    if n_src < spec.get("min_fuentes", 1):
        raise RuntimeError(f"Dato {best_key} solo en {n_src} fuente(s)")
    import pandas as pd
    s = pd.Series([best_key[2]], index=pd.DatetimeIndex([pd.Timestamp(best_key[0], best_key[1], 1)]))
    res = SourceResult(s, f"Titulares ({n_src} fuente{'s' if n_src > 1 else ''})", "https://news.google.com",
                       spec["consulta"], verified=False)
    res.notes.append("Dato extraído de titulares: verificar contra el comunicado oficial")
    return res
