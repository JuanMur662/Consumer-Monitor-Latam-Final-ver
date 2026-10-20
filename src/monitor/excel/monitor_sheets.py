"""
Capa "de negocio" sobre el XML: ubica en cada hoja la tabla de input
(Serie Mensual Histórica), el heatmap y los gráficos, y los actualiza.

La ubicación es por contenido (celda "Fecha" en la columna B, celda "Mes" para
el heatmap), así que si mueves bloques dentro de la hoja el código los sigue
encontrando. Cada corrida además deja rangos con nombre en el Administrador de
nombres (p. ej. ActEco_Datos, ActEco_Heatmap).
"""
from __future__ import annotations

import datetime as dt
import math
import unicodedata
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import pandas as pd
from lxml import etree

from .xlsx_book import NS, Sheet, XlsxBook, col_to_idx, from_serial, idx_to_col, quote_sheet, split_ref, to_serial

C = "{%s}" % NS["c"]

MESES_ES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

GREEN = "FF008000"    # dato oficial descargado (o confirmado) automáticamente
ORANGE = "FFC55A11"   # dato manual no verificado / extraído de noticias
YELLOW = "FFFFF2CC"   # celda pendiente (mes sin publicar)


def norm_label(s) -> str:
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return " ".join(s.upper().split())


def month_key(d) -> Tuple[int, int]:
    return (d.year, d.month)


def label_es(d) -> str:
    return f"{MESES_ES[d.month - 1]}-{str(d.year)[2:]}"


def add_months(d: dt.date, n: int) -> dt.date:
    y, m = d.year, d.month + n
    while m > 12:
        y, m = y + 1, m - 12
    while m < 1:
        y, m = y - 1, m + 12
    return dt.date(y, m, 1)


def month_end(d: dt.date) -> dt.date:
    nxt = add_months(dt.date(d.year, d.month, 1), 1)
    return nxt - dt.timedelta(days=1)


def _is_date_serial(v) -> bool:
    return isinstance(v, float) and 20000 < v < 80000


@dataclass
class Change:
    sheet: str
    column: str
    cell: str
    month: str
    old: Optional[float]
    new: Optional[float]
    source: str
    kind: str  # nuevo | revision | confirmado


@dataclass
class DataTable:
    book: XlsxBook
    sheet: Sheet
    header_row: int
    date_col: str
    columns: Dict[str, str]                       # etiqueta normalizada -> letra
    labels: Dict[str, str]                        # etiqueta normalizada -> etiqueta original
    rows: List[Tuple[int, dt.date]] = field(default_factory=list)
    month_end_convention: bool = False

    # -- localización ----------------------------------------------------------
    @classmethod
    def locate(cls, book: XlsxBook, sheet_name: str, date_col: str = "B") -> "DataTable":
        sh = book.sheet(sheet_name)
        header = None
        for r in range(1, min(sh.max_row(), 400) + 1):
            v = sh.get(f"{date_col}{r}")
            if isinstance(v, str) and norm_label(v) == "FECHA":
                nxt = sh.get(f"{date_col}{r + 1}")
                if _is_date_serial(nxt) or sh.has_formula(f"{date_col}{r + 1}"):
                    header = r
                    break
        if header is None:
            raise ValueError(f"No encontré la tabla de input (celda 'Fecha') en la hoja {sheet_name}")
        cols, labels = {}, {}
        ci = col_to_idx(date_col) + 1
        while True:
            col = idx_to_col(ci)
            v = sh.get(f"{col}{header}")
            if not isinstance(v, str) or not v.strip():
                break
            cols[norm_label(v)] = col
            labels[norm_label(v)] = v.strip()
            ci += 1
        t = cls(book, sh, header, date_col, cols, labels)
        t._scan_rows()
        return t

    def _scan_rows(self):
        self.rows = []
        r = self.header_row + 1
        prev = None
        while True:
            v = self.sheet.get(f"{self.date_col}{r}")
            if _is_date_serial(v):
                d = from_serial(v).date()
            elif self.sheet.has_formula(f"{self.date_col}{r}") and prev is not None:
                d = month_end(add_months(prev, 1))  # p. ej. =EOMONTH(B89,1) sin valor cacheado
            else:
                break
            self.rows.append((r, d))
            prev = d
            r += 1
        # convención de la hoja: primer día del mes (mayoría) o fin de mes (p. ej. 7_TPM's)
        ends = sum(1 for _, d in self.rows if d.day >= 28)
        self.month_end_convention = bool(self.rows) and ends > len(self.rows) / 2

    # -- lectura -----------------------------------------------------------------
    @property
    def first_row(self) -> int:
        return self.rows[0][0]

    @property
    def last_row(self) -> int:
        return self.rows[-1][0]

    def row_for(self, year: int, month: int) -> Optional[int]:
        for r, d in self.rows:
            if d.year == year and d.month == month:
                return r
        return None

    def read(self, label: str) -> pd.Series:
        col = self.columns[norm_label(label)]
        idx, vals = [], []
        for r, d in self.rows:
            v = self.sheet.get(f"{col}{r}")
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                idx.append(pd.Timestamp(d.year, d.month, 1))
                vals.append(float(v))
        return pd.Series(vals, index=pd.DatetimeIndex(idx), dtype=float)

    def last_data_row(self, labels: Optional[List[str]] = None) -> int:
        cols = [self.columns[norm_label(l)] for l in labels] if labels else list(self.columns.values())
        last = self.first_row
        for r, _ in self.rows:
            for c in cols:
                v = self.sheet.get(f"{c}{r}")
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    last = r
                    break
        return last

    def last_data_date(self, labels: Optional[List[str]] = None) -> Optional[dt.date]:
        r = self.last_data_row(labels)
        for rr, d in self.rows:
            if rr == r:
                return d
        return None

    # -- escritura -----------------------------------------------------------------
    def _base_style(self, col: str, r: int) -> Optional[int]:
        s = self.sheet.style(f"{col}{r}")
        if s is not None:
            return s
        for rr in range(r - 1, self.header_row, -1):
            s = self.sheet.style(f"{col}{rr}")
            if s is not None:
                return s
        return None

    def ensure_month(self, d: dt.date, pending_months_after: int = 0) -> int:
        """Garantiza que exista la fila del mes d (inserta filas al final de la tabla)."""
        target = add_months(dt.date(d.year, d.month, 1), pending_months_after)
        last_d = self.rows[-1][1]
        months_missing = (target.year - last_d.year) * 12 + (target.month - last_d.month)
        if months_missing > 0:
            at = self.last_row + 1
            src_row = self.last_row
            self.sheet.insert_rows(at, months_missing, copy_style_from=src_row)
            date_style = self.sheet.style(f"{self.date_col}{src_row}")
            cur = dt.date(last_d.year, last_d.month, 1)
            for k in range(months_missing):
                cur = add_months(cur, 1)
                dd = month_end(cur) if self.month_end_convention else cur
                self.sheet.set_date(f"{self.date_col}{at + k}", dd, date_style)
                for col in self.columns.values():
                    base = self._base_style(col, src_row)
                    st = self.book.styles.derive(base, fill_rgb=YELLOW) if base is not None else None
                    self.sheet.set_number(f"{col}{at + k}", None, st)
            self._scan_rows()
        r = self.row_for(d.year, d.month)
        assert r is not None
        return r

    def normalize_formula_dates(self):
        """Fechas fijas y con una sola convención por hoja (el original mezclaba 1-de-mes con =EOMONTH())."""
        for r, d in self.rows:
            ref = f"{self.date_col}{r}"
            dd = month_end(d) if self.month_end_convention else dt.date(d.year, d.month, 1)
            if self.sheet.has_formula(ref) or dd != d:
                self.sheet.set_date(ref, dd, self.sheet.style(ref))
        self._scan_rows()

    def write_series(self, label: str, s: pd.Series, source: str, verified: bool = True,
                     tol: float = 1e-10) -> List[Change]:
        col = self.columns[norm_label(label)]
        changes: List[Change] = []
        s = s.dropna().sort_index()
        if s.empty:
            return changes
        first_d = self.rows[0][1]
        s = s[s.index >= pd.Timestamp(first_d.year, first_d.month, 1)]
        if s.empty:
            return changes
        last = s.index.max()
        self.ensure_month(dt.date(last.year, last.month, 1))
        color = GREEN if verified else ORANGE
        for ts, val in s.items():
            r = self.row_for(ts.year, ts.month)
            if r is None:
                continue
            ref = f"{col}{r}"
            old = self.sheet.get(ref)
            old = float(old) if isinstance(old, (int, float)) and not isinstance(old, bool) else None
            base = self._base_style(col, r)
            new_style = self.book.styles.derive(base, font_rgb=color, no_fill=True) if base is not None else None
            if old is not None and abs(old - val) <= tol:
                kind = "confirmado"
            elif old is None:
                kind = "nuevo"
            else:
                kind = "revision"
            val = round(float(val), 12)  # evita ruido binario tipo 0.040230379999999996
            self.sheet.set_number(ref, val, new_style)
            if kind != "confirmado":
                changes.append(Change(self.sheet.name, self.labels[norm_label(label)], ref,
                                      f"{ts.year}-{ts.month:02d}", old, float(val), source, kind))
        return changes

    def defined_name_ref(self) -> str:
        last_col = max(self.columns.values(), key=col_to_idx)
        return f"{quote_sheet(self.sheet.name)}!${self.date_col}${self.header_row}:${last_col}${self.last_row}"


# ----------------------------------------------------------------------------
# Heatmap
# ----------------------------------------------------------------------------
@dataclass
class Heatmap:
    sheet: Sheet
    header_row: int
    rows: List[int]
    promedio_row: Optional[int]
    columns: Dict[str, str]

    @classmethod
    def locate(cls, sheet: Sheet, label_col: str = "B") -> Optional["Heatmap"]:
        header = None
        for r in range(1, min(sheet.max_row(), 80) + 1):
            v = sheet.get(f"{label_col}{r}")
            if isinstance(v, str) and norm_label(v) == "MES":
                header = r
                break
        if header is None:
            return None
        cols = {}
        ci = col_to_idx(label_col) + 1
        while True:
            col = idx_to_col(ci)
            v = sheet.get(f"{col}{header}")
            if not isinstance(v, str) or not v.strip():
                break
            cols[norm_label(v)] = col
            ci += 1
        rows, prom = [], None
        r = header + 1
        while r < header + 60:
            v = sheet.get(f"{label_col}{r}")
            if isinstance(v, str) and norm_label(v).startswith("PROMEDIO"):
                prom = r
                break
            if v is None and not sheet.has_formula(f"{label_col}{r}"):
                break
            rows.append(r)
            r += 1
        return cls(sheet, header, rows, prom, cols)

    def refresh(self, table: DataTable, avg_last: int = 12) -> Optional[dt.date]:
        """Ventana móvil: termina en el último mes con algún dato en las columnas del heatmap."""
        labels = [table.labels[k] for k in self.columns if k in table.columns]
        end = table.last_data_date(labels)
        if end is None:
            return None
        n = len(self.rows)
        start = add_months(dt.date(end.year, end.month, 1), -(n - 1))
        label_style = self.sheet.style(f"B{self.rows[0]}")
        # estilo homogéneo por columna (el original tenía filas con 1 y 2 decimales mezclados)
        col_style = {k: self.sheet.style(f"{c}{self.rows[0]}") for k, c in self.columns.items()}
        for i, hr in enumerate(self.rows):
            m = add_months(start, i)
            st = self.sheet.style(f"B{hr}") or label_style
            self.sheet.set_text(f"B{hr}", label_es(m), st)
            dr = table.row_for(m.year, m.month)
            for key, hcol in self.columns.items():
                tcol = table.columns.get(key)
                cst = col_style.get(key) or self.sheet.style(f"{hcol}{hr}")
                if tcol is None or dr is None:
                    self.sheet.set_formula(f"{hcol}{hr}", '""', "", cst)
                    continue
                v = table.sheet.get(f"{tcol}{dr}")
                cached = float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else ""
                self.sheet.set_formula(f"{hcol}{hr}", f'IFERROR(IF({tcol}{dr}="","",{tcol}{dr}),"")', cached, cst)
        if self.promedio_row:
            k = min(avg_last, n)
            a, b = self.rows[n - k], self.rows[-1]
            for key, hcol in self.columns.items():
                cst = self.sheet.style(f"{hcol}{self.promedio_row}")
                vals = []
                for hr in self.rows[n - k:]:
                    v = self.sheet.get(f"{hcol}{hr}")
                    if isinstance(v, float):
                        vals.append(v)
                cached = sum(vals) / len(vals) if vals else ""
                self.sheet.set_formula(f"{hcol}{self.promedio_row}", f'IFERROR(AVERAGE({hcol}{a}:{hcol}{b}),"")',
                                       cached, cst)
        return end

    def defined_name_ref(self) -> str:
        last_col = max(self.columns.values(), key=col_to_idx)
        last = self.promedio_row or self.rows[-1]
        return f"{quote_sheet(self.sheet.name)}!$B${self.header_row}:${last_col}${last}"


# ----------------------------------------------------------------------------
# Gráficos
# ----------------------------------------------------------------------------
def _parse_chart_ref(text: str):
    """'Hoja'!$B$24:$B$86 -> (hoja, colA, r1, colB, r2)"""
    if "!" not in text:
        return None
    sh, rng = text.rsplit("!", 1)
    if sh.startswith("'"):
        sh = sh[1:-1].replace("''", "'")
    if ":" not in rng:
        return None
    a, b = rng.split(":")
    try:
        ca, ra = split_ref(a)
        cb, rb = split_ref(b)
    except ValueError:
        return None
    return sh, ca, ra, cb, rb


def update_charts(book: XlsxBook, table: DataTable, end_row: Optional[int] = None) -> int:
    """Ajusta el rango de cada serie de los gráficos de la hoja hasta end_row y reconstruye cachés."""
    end_row = end_row or table.last_data_row()
    n_updated = 0
    for part in book.charts_of(table.sheet.name):
        ch = book.chart(part)
        for f in ch.iter(C + "f"):
            parsed = _parse_chart_ref(f.text or "")
            if not parsed:
                continue
            sh, ca, ra, cb, rb = parsed
            if sh != table.sheet.name or ca != cb or ra != table.first_row:
                continue
            new_ref = f"{quote_sheet(sh)}!${ca}${ra}:${cb}${end_row}"
            f.text = new_ref
            n_updated += 1
            ref_el = f.getparent()  # numRef o strRef
            cache = ref_el.find(C + "numCache")
            if cache is None:
                cache = ref_el.find(C + "strCache")
            if cache is None:
                continue
            is_num = cache.tag == C + "numCache"
            for pt in cache.findall(C + "pt"):
                cache.remove(pt)
            ptc = cache.find(C + "ptCount")
            if ptc is None:
                ptc = etree.SubElement(cache, C + "ptCount")
            npts = end_row - ra + 1
            ptc.set("val", str(npts))
            anchor = ptc
            for i, r in enumerate(range(ra, end_row + 1)):
                v = table.sheet.get(f"{ca}{r}")
                if v is None or v == "":
                    continue
                if is_num and not isinstance(v, (int, float)):
                    continue
                pt = etree.Element(C + "pt", idx=str(i))
                etree.SubElement(pt, C + "v").text = (repr(float(v)) if is_num else str(v))
                anchor.addnext(pt)
                anchor = pt
    return n_updated
