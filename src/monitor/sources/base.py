"""Utilidades comunes a todas las fuentes."""
from __future__ import annotations

import datetime as dt
import math
import re
import unicodedata
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, Optional

import numpy as np
import pandas as pd

MONTHS = {
    # español
    "ene": 1, "enero": 1, "feb": 2, "febrero": 2, "mar": 3, "marzo": 3, "abr": 4, "abril": 4,
    "may": 5, "mayo": 5, "jun": 6, "junio": 6, "jul": 7, "julio": 7, "ago": 8, "agosto": 8,
    "sep": 9, "set": 9, "sept": 9, "septiembre": 9, "setiembre": 9, "oct": 10, "octubre": 10,
    "nov": 11, "noviembre": 11, "dic": 12, "diciembre": 12,
    # portugués
    "fev": 2, "fevereiro": 2, "marco": 3, "abril_pt": 4, "mai": 5, "maio": 5, "junho": 6, "julho": 7,
    "outubro": 10, "out": 10, "novembro": 11, "dez": 12, "dezembro": 12, "janeiro": 1, "jan": 1,
    "setembro": 9,
    # inglés
    "january": 1, "february": 2, "march": 3, "april": 4, "june": 6, "july": 7, "august": 8, "aug": 8,
    "september": 9, "october": 10, "november": 11, "december": 12, "dec": 12, "apr": 4,
}


def strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFD", str(s))
    return "".join(ch for ch in s if unicodedata.category(ch) != "Mn")


def month_from_name(s: str) -> Optional[int]:
    k = strip_accents(s).lower().strip(" .")
    return MONTHS.get(k)


_PERIOD_PATTERNS = [
    # 2026-08 / 2026/08 / 2026M08 / 202608
    (re.compile(r"^\s*(\d{4})\s*[-/M]?\s*(\d{1,2})\s*$", re.I), lambda m: (int(m.group(1)), int(m.group(2)))),
    # 08/2026
    (re.compile(r"^\s*(\d{1,2})\s*[-/.]\s*(\d{4})\s*$"), lambda m: (int(m.group(2)), int(m.group(1)))),
    # 01/08/2026 (dd/mm/yyyy)
    (re.compile(r"^\s*(\d{1,2})[/-](\d{1,2})[/-](\d{4})\s*$"), lambda m: (int(m.group(3)), int(m.group(2)))),
    # 2026-08-01
    (re.compile(r"^\s*(\d{4})-(\d{1,2})-(\d{1,2})"), lambda m: (int(m.group(1)), int(m.group(2)))),
]
_NAMED = re.compile(r"^\s*([A-Za-zÁÉÍÓÚáéíóúñÑç]+)\.?\s*[-/ .]?\s*(?:de\s+)?(\d{2}|\d{4})\s*$")
_NAMED_REV = re.compile(r"^\s*(\d{4})\s*[-/ .]?\s*([A-Za-zÁÉÍÓÚáéíóúñÑç]+)\.?\s*$")


def parse_period(v) -> Optional[pd.Timestamp]:
    """Convierte casi cualquier representación de mes en Timestamp (primer día del mes)."""
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if isinstance(v, (pd.Timestamp, dt.datetime, dt.date)):
        return pd.Timestamp(v.year, v.month, 1)
    if isinstance(v, (int, np.integer)) and 190001 <= int(v) <= 210012:
        y, m = divmod(int(v), 100)
        return pd.Timestamp(y, m, 1) if 1 <= m <= 12 else None
    s = str(v).strip()
    if not s:
        return None
    for rx, fn in _PERIOD_PATTERNS:
        m = rx.match(s)
        if m:
            y, mo = fn(m)
            if 1900 < y < 2100 and 1 <= mo <= 12:
                return pd.Timestamp(y, mo, 1)
    m = _NAMED.match(s)
    if m:
        mo = month_from_name(m.group(1))
        y = int(m.group(2))
        y = y + 2000 if y < 100 else y
        if mo:
            return pd.Timestamp(y, mo, 1)
    m = _NAMED_REV.match(s)
    if m:
        mo = month_from_name(m.group(2))
        if mo:
            return pd.Timestamp(int(m.group(1)), mo, 1)
    return None


def parse_number(v, decimal: str = "auto") -> Optional[float]:
    if v is None:
        return None
    if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
        f = float(v)
        return None if math.isnan(f) else f
    s = str(v).strip().replace("\xa0", "").replace(" ", "")
    if s in ("", "-", "--", "n.d.", "nd", "N/D", "N/E", "NaN", "nan", "NC", "...", "ND", "n.d"):
        return None
    s = s.replace("%", "")
    if decimal == ",":
        s = s.replace(".", "").replace(",", ".")
    elif decimal == ".":
        s = s.replace(",", "")
    else:  # auto
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s:
            s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def to_monthly(points, how: str = "last") -> pd.Series:
    """points: iterable de (fecha, valor). Devuelve serie mensual (índice = primer día del mes)."""
    rows = [(pd.Timestamp(d), v) for d, v in points if d is not None and v is not None]
    if not rows:
        return pd.Series(dtype=float)
    s = pd.Series([v for _, v in rows], index=pd.DatetimeIndex([d for d, _ in rows])).sort_index()
    s = s[~s.index.duplicated(keep="last")]
    g = s.groupby([s.index.year, s.index.month])
    agg = g.last() if how == "last" else g.mean()
    idx = pd.DatetimeIndex([pd.Timestamp(y, m, 1) for y, m in agg.index])
    return pd.Series(agg.values, index=idx, dtype=float)


@dataclass
class SourceResult:
    series: pd.Series                         # mensual, índice = primer día del mes, ya transformada
    source: str                               # etiqueta corta, p. ej. "BCB SGS 24363"
    url: str = ""
    title: str = ""
    notes: list = field(default_factory=list)
    verified: bool = True                     # False -> se pinta naranja (noticias / manual sin verificar)
    raw: Optional[pd.Series] = None           # antes de transformar
    alternatives: list = field(default_factory=list)   # [(etiqueta, serie_cruda)] otras filas candidatas del archivo


def apply_transform(s: pd.Series, transform: Optional[str]) -> pd.Series:
    """Transformaciones declaradas en config/series.yaml."""
    s = s.sort_index().astype(float)
    if not transform or transform == "nivel":
        return s
    if transform == "pct":          # viene en %, el Excel guarda fracción
        return s / 100.0
    if transform == "yoy":          # desde índice/nivel
        full = s.asfreq("MS")
        return (full / full.shift(12) - 1.0).dropna()
    if transform == "auto_yoy":     # índice si la mediana es alta, si no ya es variación en %
        return apply_transform(s, "yoy" if s.median() > 30 else "pct")
    if transform.startswith("formula:"):
        expr = transform.split(":", 1)[1]
        return s.apply(lambda x: float(eval(expr, {"__builtins__": {}}, {"x": x})))  # noqa: S307 (expresión de config)
    raise ValueError(f"Transformación desconocida: {transform}")


Fetcher = Callable[[Dict[str, Any], Any], SourceResult]
REGISTRY: Dict[str, Fetcher] = {}


def register(name: str):
    def deco(fn: Fetcher):
        REGISTRY[name] = fn
        return fn
    return deco
