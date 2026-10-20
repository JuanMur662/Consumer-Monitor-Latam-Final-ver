"""
Lector genérico de series dentro de archivos Excel publicados por institutos
(DANE, INE Chile, CNI, Fedesarrollo...), cuyos formatos cambian con frecuencia.

Estrategia: se buscan, en cada hoja, tres disposiciones típicas
  1. "ancha": una fila de encabezado con meses (o dos filas: años + meses) y la serie en una fila
  2. "larga": una columna con meses y la serie en otra columna (por encabezado)
  3. "matriz mes x año": meses en la primera columna y años como encabezado (típico IPC DANE)
y se devuelve la coincidencia más larga/reciente para la etiqueta pedida.
El resultado SIEMPRE pasa luego por la validación contra el histórico del Excel.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import List, Optional

import pandas as pd

from .base import month_from_name, parse_number, parse_period, strip_accents, to_monthly


@dataclass
class Candidate:
    sheet: str
    layout: str
    label: str
    series: pd.Series

    @property
    def score(self):
        s = self.series.dropna()
        return (s.index.max() if not s.empty else pd.Timestamp(1900, 1, 1), len(s))


def read_all_sheets(content: bytes, ext: str) -> dict:
    engine = "xlrd" if ext == "xls" else "openpyxl"
    return pd.read_excel(io.BytesIO(content), sheet_name=None, header=None, engine=engine)


def _norm(v) -> str:
    return strip_accents(str(v)).lower().strip() if v is not None and str(v) != "nan" else ""


def _period_row(df: pd.DataFrame, r: int) -> dict:
    out = {}
    for c in range(df.shape[1]):
        p = parse_period(df.iat[r, c])
        if p is not None:
            out[c] = p
    return out


def _two_row_periods(df: pd.DataFrame, r: int) -> dict:
    """Fila r = años (con celdas combinadas -> forward fill), fila r+1 = meses."""
    if r + 1 >= df.shape[0]:
        return {}
    years, cur = {}, None
    for c in range(df.shape[1]):
        v = df.iat[r, c]
        n = parse_number(v, ".")
        if n is not None and 1990 <= n <= 2100 and float(n).is_integer():
            cur = int(n)
        else:
            m = re.search(r"(19|20)\d{2}", str(v)) if v is not None else None
            if m and len(str(v)) <= 12:
                cur = int(m.group(0))
        years[c] = cur
    out = {}
    for c in range(df.shape[1]):
        v = df.iat[r + 1, c]
        mo = month_from_name(str(v).split("*")[0]) if isinstance(v, str) else None
        if mo is None:
            n = parse_number(v, ".")
            if n is not None and float(n).is_integer() and 1 <= n <= 12 and years.get(c):
                mo = int(n)
        if mo and years.get(c):
            out[c] = pd.Timestamp(years[c], mo, 1)
    return out


def find_series(sheets: dict, label_regex: str, sheet_regex: Optional[str] = None,
                exclude_regex: Optional[str] = None, max_scan_rows: int = 400) -> List[Candidate]:
    rx = re.compile(strip_accents(label_regex), re.I)
    srx = re.compile(strip_accents(sheet_regex), re.I) if sheet_regex else None
    xrx = re.compile(strip_accents(exclude_regex), re.I) if exclude_regex else None
    cands: List[Candidate] = []
    for sname, df in sheets.items():
        if srx and not srx.search(strip_accents(sname)):
            continue
        df = df.iloc[:max_scan_rows, :400]
        nrows, ncols = df.shape
        # ---- 1. disposición ancha
        for r in range(min(nrows, 60)):
            per = _period_row(df, r)
            if len(per) < 6:
                per = _two_row_periods(df, r)
                hdr_rows = 2
            else:
                hdr_rows = 1
            if len(per) < 6:
                continue
            first_col = min(per)
            for rr in range(r + hdr_rows, nrows):
                label = " ".join(_norm(df.iat[rr, c]) for c in range(min(first_col, 4)) if _norm(df.iat[rr, c]))
                if not label or not rx.search(label) or (xrx and xrx.search(label)):
                    continue
                pts = [(p, parse_number(df.iat[rr, c])) for c, p in per.items()]
                s = to_monthly(pts)
                if len(s.dropna()) >= 6:
                    cands.append(Candidate(sname, "ancha", label, s))
            break
        # ---- 2. disposición larga (columna de fechas)
        for c in range(min(ncols, 6)):
            col_per = [(r, parse_period(df.iat[r, c])) for r in range(nrows)]
            col_per = [(r, p) for r, p in col_per if p is not None]
            if len(col_per) < 12:
                continue
            first_r = col_per[0][0]
            # combinar años en col c y meses en col c+1 si hiciera falta se cubre en (3)
            for vc in range(ncols):
                if vc == c:
                    continue
                head = " ".join(_norm(df.iat[h, vc]) for h in range(max(0, first_r - 6), first_r) if _norm(df.iat[h, vc]))
                if not head or not rx.search(head) or (xrx and xrx.search(head)):
                    continue
                pts = [(p, parse_number(df.iat[r, vc])) for r, p in col_per]
                s = to_monthly(pts)
                if len(s.dropna()) >= 12:
                    cands.append(Candidate(sname, "larga", head, s))
        # ---- 2b. año en una columna y mes en la siguiente
        for c in range(min(ncols, 5)):
            pts_idx = []
            cur_year = None
            for r in range(nrows):
                yv = parse_number(df.iat[r, c], ".")
                if yv is not None and 1990 <= yv <= 2100 and float(yv).is_integer():
                    cur_year = int(yv)
                mo = month_from_name(str(df.iat[r, c + 1])) if c + 1 < ncols and isinstance(df.iat[r, c + 1], str) else None
                if cur_year and mo:
                    pts_idx.append((r, pd.Timestamp(cur_year, mo, 1)))
            if len(pts_idx) < 12:
                continue
            first_r = pts_idx[0][0]
            for vc in range(c + 2, ncols):
                head = " ".join(_norm(df.iat[h, vc]) for h in range(max(0, first_r - 6), first_r) if _norm(df.iat[h, vc]))
                if not head or not rx.search(head) or (xrx and xrx.search(head)):
                    continue
                s = to_monthly([(p, parse_number(df.iat[r, vc])) for r, p in pts_idx])
                if len(s.dropna()) >= 12:
                    cands.append(Candidate(sname, "anio_mes", head, s))
        # ---- 3. matriz mes x año (solo si la hoja o su título coinciden con la etiqueta)
        sheet_text = " ".join(_norm(df.iat[r, c]) for r in range(min(nrows, 12)) for c in range(min(ncols, 6)))
        if rx.search(strip_accents(sname).lower()) or rx.search(sheet_text):
            for r in range(min(nrows, 40)):
                years = {}
                for c in range(ncols):
                    n = parse_number(df.iat[r, c], ".")
                    if n is not None and 1990 <= n <= 2100 and float(n).is_integer():
                        years[c] = int(n)
                if len(years) < 4:
                    continue
                pts = []
                for rr in range(r + 1, min(nrows, r + 20)):
                    mo = None
                    for c in range(min(ncols, 3)):
                        if isinstance(df.iat[rr, c], str):
                            mo = month_from_name(df.iat[rr, c])
                            if mo:
                                break
                    if not mo:
                        continue
                    for c, y in years.items():
                        pts.append((pd.Timestamp(y, mo, 1), parse_number(df.iat[rr, c])))
                s = to_monthly(pts)
                if len(s.dropna()) >= 12:
                    cands.append(Candidate(sname, "matriz", f"{sname} (mes x año)", s))
                break
    cands.sort(key=lambda c: c.score, reverse=True)
    return cands


def describe(sheets: dict, max_rows: int = 12) -> str:
    """Resumen legible de un archivo (para el modo diagnóstico)."""
    out = []
    for sname, df in sheets.items():
        out.append(f"--- hoja '{sname}' ({df.shape[0]}x{df.shape[1]})")
        for r in range(min(max_rows, df.shape[0])):
            vals = [str(v)[:18] for v in df.iloc[r, :10].tolist() if str(v) != "nan"]
            if vals:
                out.append("   " + " | ".join(vals))
    return "\n".join(out)
