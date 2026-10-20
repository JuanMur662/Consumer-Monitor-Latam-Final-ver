"""
Edición "quirúrgica" de un .xlsx a nivel XML.

Por qué no openpyxl: openpyxl reescribe el libro completo y pierde formato de
gráficos (extensiones c14/c15, estilos de series, etiquetas), formatos
condicionales con extLst, etc. Aquí solo se tocan las celdas, rangos y cachés
que el pipeline necesita; todo lo demás del archivo se copia byte a byte.

Funciones principales
- leer / escribir valores, fechas, textos y fórmulas (con valor cacheado)
- insertar filas desplazando celdas, merges, formatos condicionales,
  hipervínculos, anclas de dibujos, nombres definidos y fórmulas
- actualizar rangos y cachés de gráficos
- derivar estilos (color de fuente / relleno) sin alterar el formato numérico
- crear hojas nuevas (Noticias, Log) con estilos propios
- guardar con recálculo completo al abrir (fullCalcOnLoad)
"""
from __future__ import annotations

import copy
import datetime as dt
import math
import posixpath
import re
import zipfile
from typing import Dict, Iterable, List, Optional, Tuple

from lxml import etree

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
}
M = "{%s}" % NS["m"]
R_ID = "{%s}id" % NS["r"]
REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
REL_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

EXCEL_EPOCH = dt.datetime(1899, 12, 30)

# ----------------------------------------------------------------------------
# Utilidades de referencias A1
# ----------------------------------------------------------------------------
_CELL_RE = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)$")


def col_to_idx(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def idx_to_col(idx: int) -> str:
    s = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def split_ref(ref: str) -> Tuple[str, int]:
    m = _CELL_RE.match(ref)
    if not m:
        raise ValueError(f"Referencia inválida: {ref}")
    return m.group(1), int(m.group(2))


def to_serial(d: dt.date) -> float:
    if isinstance(d, dt.datetime):
        delta = d - EXCEL_EPOCH
    else:
        delta = dt.datetime(d.year, d.month, d.day) - EXCEL_EPOCH
    return delta.days + delta.seconds / 86400.0


def from_serial(v: float) -> dt.datetime:
    return EXCEL_EPOCH + dt.timedelta(days=float(v))


def quote_sheet(name: str) -> str:
    if re.match(r"^[A-Za-z_][A-Za-z0-9_.]*$", name):
        return name
    return "'" + name.replace("'", "''") + "'"


def fmt_number(v: float) -> str:
    if isinstance(v, bool):
        return "1" if v else "0"
    if float(v).is_integer() and abs(v) < 1e15:
        return str(int(v))
    return repr(float(v))


# Desplazamiento de referencias dentro de fórmulas -----------------------------
_FORMULA_REF_RE = re.compile(
    r"(?P<sheet>(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+)!)?"
    r"(?P<c1>\$?[A-Z]{1,3})(?P<r1d>\$?)(?P<r1>\d+)"
    r"(?::(?P<c2>\$?[A-Z]{1,3})(?P<r2d>\$?)(?P<r2>\d+))?"
    r"(?![A-Za-z0-9_(])"
)


def _unquote_sheet(s: Optional[str]) -> Optional[str]:
    if not s:
        return None
    s = s[:-1]  # quita "!"
    if s.startswith("'"):
        s = s[1:-1].replace("''", "'")
    return s


def shift_formula_rows(formula: str, target_sheet: str, host_sheet: str, at_row: int, n: int) -> str:
    """Desplaza n filas las referencias a `target_sheet` con fila >= at_row.

    `host_sheet` es la hoja donde vive la fórmula (para referencias sin hoja).
    Respeta textos entre comillas dobles.
    """
    out = []
    parts = re.split(r'("(?:[^"]|"")*")', formula)
    for part in parts:
        if part.startswith('"'):
            out.append(part)
            continue

        def repl(m: re.Match) -> str:
            sheet = _unquote_sheet(m.group("sheet")) or host_sheet
            if sheet != target_sheet:
                return m.group(0)
            r1 = int(m.group("r1"))
            r1n = r1 + n if r1 >= at_row else r1
            s = (m.group("sheet") or "") + m.group("c1") + m.group("r1d") + str(r1n)
            if m.group("c2"):
                r2 = int(m.group("r2"))
                r2n = r2 + n if r2 >= at_row else r2
                s += ":" + m.group("c2") + m.group("r2d") + str(r2n)
            return s

        out.append(_FORMULA_REF_RE.sub(repl, part))
    return "".join(out)


def shift_sqref(sqref: str, at_row: int, n: int) -> str:
    items = []
    for rng in sqref.split():
        if ":" in rng:
            a, b = rng.split(":")
            ca, ra = split_ref(a)
            cb, rb = split_ref(b)
            ra = ra + n if ra >= at_row else ra
            rb = rb + n if rb >= at_row else rb
            items.append(f"{ca}{ra}:{cb}{rb}")
        else:
            c, r = split_ref(rng)
            r = r + n if r >= at_row else r
            items.append(f"{c}{r}")
    return " ".join(items)


# ----------------------------------------------------------------------------
# Estilos
# ----------------------------------------------------------------------------
class StyleBook:
    """Acceso mínimo a styles.xml para derivar estilos sin romper los existentes."""

    def __init__(self, xml: bytes):
        self.root = etree.fromstring(xml)
        self.fonts = self.root.find("m:fonts", NS)
        self.fills = self.root.find("m:fills", NS)
        self.borders = self.root.find("m:borders", NS)
        self.cellXfs = self.root.find("m:cellXfs", NS)
        self.numFmts = self.root.find("m:numFmts", NS)
        self._cache: Dict[tuple, int] = {}

    # -- helpers -------------------------------------------------------------
    @classmethod
    def _sig(cls, el) -> tuple:
        """Firma independiente del prefijo de namespace (un elemento nuevo se serializa como ns0:...)."""
        return (etree.QName(el).localname, tuple(sorted(el.attrib.items())), (el.text or "").strip(),
                tuple(cls._sig(ch) for ch in el))

    def _append(self, parent, el) -> int:
        """Agrega el elemento salvo que ya exista uno idéntico (evita que styles.xml crezca en cada corrida)."""
        sig = self._sig(el)
        children = parent.findall("*")
        for i, ch in enumerate(children):
            if self._sig(ch) == sig:
                return i
        parent.append(el)
        cnt = len(children) + 1
        parent.set("count", str(cnt))
        return cnt - 1

    def xf(self, idx: int):
        return self.cellXfs.findall("m:xf", NS)[idx]

    def num_format_code(self, idx: int) -> str:
        xf = self.xf(idx)
        fid = int(xf.get("numFmtId", "0"))
        if self.numFmts is not None:
            for nf in self.numFmts.findall("m:numFmt", NS):
                if int(nf.get("numFmtId")) == fid:
                    return nf.get("formatCode")
        builtin = {0: "General", 9: "0%", 10: "0.00%", 14: "m/d/yyyy", 17: "mmm-yy"}
        return builtin.get(fid, "General")

    def num_fmt_id(self, code: str) -> int:
        builtin = {"General": 0, "0": 1, "0.00": 2, "0%": 9, "0.00%": 10, "d-mmm-yy": 15, "mmm-yy": 17}
        if code in builtin:
            return builtin[code]
        if self.numFmts is None:
            self.numFmts = etree.Element(M + "numFmts")
            self.root.insert(0, self.numFmts)
        ids = [int(n.get("numFmtId")) for n in self.numFmts.findall("m:numFmt", NS)]
        for n in self.numFmts.findall("m:numFmt", NS):
            if n.get("formatCode") == code:
                return int(n.get("numFmtId"))
        new_id = max([163] + ids) + 1
        el = etree.SubElement(self.numFmts, M + "numFmt")
        el.set("numFmtId", str(new_id))
        el.set("formatCode", code)
        self.numFmts.set("count", str(len(self.numFmts.findall("m:numFmt", NS))))
        return new_id

    def _font_with_color(self, font_id: int, rgb: str) -> int:
        base = self.fonts.findall("m:font", NS)[font_id]
        new = copy.deepcopy(base)
        col = new.find("m:color", NS)
        if col is None:
            col = etree.Element(M + "color")
            # el orden importa poco para Excel, pero colocamos color tras sz si existe
            sz = new.find("m:sz", NS)
            if sz is not None:
                sz.addnext(col)
            else:
                new.insert(0, col)
        for a in list(col.attrib):
            del col.attrib[a]
        col.set("rgb", rgb)
        return self._append(self.fonts, new)

    def _solid_fill(self, rgb: str) -> int:
        for i, f in enumerate(self.fills.findall("m:fill", NS)):
            p = f.find("m:patternFill", NS)
            if p is not None and p.get("patternType") == "solid":
                fg = p.find("m:fgColor", NS)
                if fg is not None and fg.get("rgb") == rgb:
                    return i
        fill = etree.Element(M + "fill")
        p = etree.SubElement(fill, M + "patternFill", patternType="solid")
        etree.SubElement(p, M + "fgColor", rgb=rgb)
        etree.SubElement(p, M + "bgColor", indexed="64")
        return self._append(self.fills, fill)

    def derive(self, base_idx: int, font_rgb: Optional[str] = None, fill_rgb: Optional[str] = None,
               no_fill: bool = False, num_fmt: Optional[str] = None) -> int:
        """Devuelve un índice de estilo igual a base_idx pero con cambios puntuales."""
        key = (base_idx, font_rgb, fill_rgb, no_fill, num_fmt)
        if key in self._cache:
            return self._cache[key]
        base = self.xf(base_idx)
        new = copy.deepcopy(base)
        if font_rgb:
            fid = self._font_with_color(int(base.get("fontId", "0")), font_rgb)
            new.set("fontId", str(fid))
            new.set("applyFont", "1")
        if no_fill:
            new.set("fillId", "0")
        elif fill_rgb:
            new.set("fillId", str(self._solid_fill(fill_rgb)))
            new.set("applyFill", "1")
        if num_fmt:
            new.set("numFmtId", str(self.num_fmt_id(num_fmt)))
            new.set("applyNumberFormat", "1")
        idx = self._append(self.cellXfs, new)
        self._cache[key] = idx
        return idx

    def new_style(self, bold=False, italic=False, size=10, color="FF000000", fill_rgb=None,
                  num_fmt="General", wrap=False, halign=None, valign="top", underline=False,
                  border=False, font_name="Arial") -> int:
        key = ("new", bold, italic, size, color, fill_rgb, num_fmt, wrap, halign, valign, underline, border, font_name)
        if key in self._cache:
            return self._cache[key]
        font = etree.Element(M + "font")
        if bold:
            etree.SubElement(font, M + "b")
        if italic:
            etree.SubElement(font, M + "i")
        if underline:
            etree.SubElement(font, M + "u")
        etree.SubElement(font, M + "sz", val=str(size))
        etree.SubElement(font, M + "color", rgb=color)
        etree.SubElement(font, M + "name", val=font_name)
        etree.SubElement(font, M + "family", val="2")
        font_id = self._append(self.fonts, font)
        fill_id = self._solid_fill(fill_rgb) if fill_rgb else 0
        border_id = 0
        if border:
            b = etree.Element(M + "border")
            for side in ("left", "right", "top", "bottom"):
                s = etree.SubElement(b, M + side, style="thin")
                etree.SubElement(s, M + "color", rgb="FFBFBFBF")
            etree.SubElement(b, M + "diagonal")
            border_id = self._append(self.borders, b)
        xf = etree.Element(M + "xf", numFmtId=str(self.num_fmt_id(num_fmt)), fontId=str(font_id),
                           fillId=str(fill_id), borderId=str(border_id), xfId="0",
                           applyFont="1", applyNumberFormat="1")
        if fill_id:
            xf.set("applyFill", "1")
        if border_id:
            xf.set("applyBorder", "1")
        al = etree.SubElement(xf, M + "alignment")
        if wrap:
            al.set("wrapText", "1")
        if halign:
            al.set("horizontal", halign)
        if valign:
            al.set("vertical", valign)
        xf.set("applyAlignment", "1")
        idx = self._append(self.cellXfs, xf)
        self._cache[key] = idx
        return idx

    def tobytes(self) -> bytes:
        return etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ----------------------------------------------------------------------------
# Hoja
# ----------------------------------------------------------------------------
class Sheet:
    def __init__(self, book: "XlsxBook", name: str, path: str, xml: bytes):
        self.book = book
        self.name = name
        self.path = path
        self.root = etree.fromstring(xml)
        self.sheetData = self.root.find("m:sheetData", NS)
        self._rows: Dict[int, etree._Element] = {}
        self._reindex()

    def _reindex(self):
        self._rows = {int(r.get("r")): r for r in self.sheetData.findall("m:row", NS)}

    # -- filas y celdas --------------------------------------------------------
    def row_el(self, r: int, create: bool = True):
        el = self._rows.get(r)
        if el is None and create:
            el = etree.Element(M + "row", r=str(r))
            # insertar en orden
            after = None
            for rr in sorted(self._rows):
                if rr < r:
                    after = self._rows[rr]
                else:
                    break
            if after is None:
                self.sheetData.insert(0, el)
            else:
                after.addnext(el)
            self._rows[r] = el
        return el

    def cell_el(self, ref: str, create: bool = False):
        col, r = split_ref(ref)
        row = self.row_el(r, create=create)
        if row is None:
            return None
        for c in row.findall("m:c", NS):
            if c.get("r") == f"{col}{r}":
                return c
        if not create:
            return None
        c = etree.Element(M + "c", r=f"{col}{r}")
        ci = col_to_idx(col)
        placed = False
        for other in row.findall("m:c", NS):
            oc, _ = split_ref(other.get("r"))
            if col_to_idx(oc) > ci:
                other.addprevious(c)
                placed = True
                break
        if not placed:
            row.append(c)
        # actualizar spans no es obligatorio; se elimina para evitar inconsistencias
        if "spans" in row.attrib:
            del row.attrib["spans"]
        return c

    def get(self, ref: str):
        """Valor de la celda: float, str, None. Para fórmulas devuelve el valor cacheado."""
        c = self.cell_el(ref)
        if c is None:
            return None
        t = c.get("t")
        v = c.find("m:v", NS)
        if t == "s":
            return self.book.shared_strings[int(v.text)] if v is not None and v.text else None
        if t == "inlineStr":
            is_ = c.find("m:is", NS)
            return "".join(is_.itertext()) if is_ is not None else None
        if v is None or v.text is None:
            return None
        if t in ("str", "e"):
            return v.text
        if t == "b":
            return bool(int(v.text))
        try:
            return float(v.text)
        except ValueError:
            return v.text

    def formula(self, ref: str) -> Optional[str]:
        c = self.cell_el(ref)
        if c is None:
            return None
        f = c.find("m:f", NS)
        return f.text if f is not None else None

    def style(self, ref: str) -> Optional[int]:
        c = self.cell_el(ref)
        if c is None or c.get("s") is None:
            return None
        return int(c.get("s"))

    def has_formula(self, ref: str) -> bool:
        c = self.cell_el(ref)
        return c is not None and c.find("m:f", NS) is not None

    def _clear(self, c):
        f = c.find("m:f", NS)
        if f is not None and f.get("t") == "shared" and f.get("ref"):
            # c es la celda "maestra" de una fórmula compartida: las celdas hijas quedarían huérfanas
            # (Excel reportaría el archivo como dañado). Se convierten en valores (su último resultado).
            si = f.get("si")
            for other in self.sheetData.iter(M + "f"):
                if other is not f and other.get("t") == "shared" and other.get("si") == si and not other.get("ref"):
                    other.getparent().remove(other)
        for child in list(c):
            c.remove(child)
        if "t" in c.attrib:
            del c.attrib["t"]

    def set_number(self, ref: str, value: Optional[float], style: Optional[int] = None):
        c = self.cell_el(ref, create=True)
        self._clear(c)
        if style is not None:
            c.set("s", str(style))
        if value is None or (isinstance(value, float) and math.isnan(value)):
            return
        etree.SubElement(c, M + "v").text = fmt_number(value)

    def set_date(self, ref: str, d: dt.date, style: Optional[int] = None):
        self.set_number(ref, to_serial(d), style)

    def set_text(self, ref: str, text: Optional[str], style: Optional[int] = None):
        c = self.cell_el(ref, create=True)
        self._clear(c)
        if style is not None:
            c.set("s", str(style))
        if text is None or text == "":
            return
        idx = self.book.add_shared_string(str(text))
        c.set("t", "s")
        etree.SubElement(c, M + "v").text = str(idx)

    def set_formula(self, ref: str, formula: str, cached=None, style: Optional[int] = None):
        c = self.cell_el(ref, create=True)
        self._clear(c)
        if style is not None:
            c.set("s", str(style))
        etree.SubElement(c, M + "f").text = formula.lstrip("=")
        if cached is None or cached == "" or (isinstance(cached, float) and math.isnan(cached)):
            c.set("t", "str")
            etree.SubElement(c, M + "v").text = ""
        elif isinstance(cached, str):
            c.set("t", "str")
            etree.SubElement(c, M + "v").text = cached
        else:
            etree.SubElement(c, M + "v").text = fmt_number(cached)

    def max_row(self) -> int:
        return max(self._rows) if self._rows else 0

    def iter_col(self, col: str, r0: int, r1: int):
        for r in range(r0, r1 + 1):
            yield r, self.get(f"{col}{r}")

    # -- inserción de filas ----------------------------------------------------
    def _shift_rows_local(self, at_row: int, n: int):
        rows = sorted(self._rows.items(), key=lambda x: -x[0])
        for r, el in rows:
            if r >= at_row:
                el.set("r", str(r + n))
                for c in el.findall("m:c", NS):
                    col, rr = split_ref(c.get("r"))
                    c.set("r", f"{col}{rr + n}")
        self._reindex()
        # merges
        mc = self.root.find("m:mergeCells", NS)
        if mc is not None:
            for m_ in mc.findall("m:mergeCell", NS):
                m_.set("ref", shift_sqref(m_.get("ref"), at_row, n))
        # formatos condicionales
        for cf in self.root.findall("m:conditionalFormatting", NS):
            cf.set("sqref", shift_sqref(cf.get("sqref"), at_row, n))
            for f in cf.iter(M + "formula"):
                if f.text:
                    f.text = shift_formula_rows(f.text, self.name, self.name, at_row, n)
        dvs = self.root.find("m:dataValidations", NS)
        if dvs is not None:
            for dv in dvs.findall("m:dataValidation", NS):
                dv.set("sqref", shift_sqref(dv.get("sqref"), at_row, n))
        hls = self.root.find("m:hyperlinks", NS)
        if hls is not None:
            for h in hls.findall("m:hyperlink", NS):
                h.set("ref", shift_sqref(h.get("ref"), at_row, n))
        dim = self.root.find("m:dimension", NS)
        if dim is not None and ":" in dim.get("ref", ""):
            dim.set("ref", shift_sqref(dim.get("ref"), at_row, n))

    def insert_rows(self, at_row: int, n: int, copy_style_from: Optional[int] = None):
        """Inserta n filas antes de `at_row` (1-based) y actualiza todo el libro."""
        if n <= 0:
            return
        self._shift_rows_local(at_row, n)
        self.book._after_rows_inserted(self, at_row, n)
        if copy_style_from is not None:
            src = self._rows.get(copy_style_from)
            for k in range(n):
                new_r = at_row + k
                row = self.row_el(new_r, create=True)
                if src is not None:
                    for a in ("ht", "customHeight", "s", "customFormat"):
                        if src.get(a) is not None:
                            row.set(a, src.get(a))
                    for c in src.findall("m:c", NS):
                        col, _ = split_ref(c.get("r"))
                        nc = self.cell_el(f"{col}{new_r}", create=True)
                        if c.get("s"):
                            nc.set("s", c.get("s"))

    def tobytes(self) -> bytes:
        return etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ----------------------------------------------------------------------------
# Libro
# ----------------------------------------------------------------------------
class XlsxBook:
    def __init__(self, path: str):
        self.path = path
        with zipfile.ZipFile(path) as z:
            self.order = [i.filename for i in z.infolist()]
            self.parts: Dict[str, bytes] = {n: z.read(n) for n in self.order}
        self.workbook = etree.fromstring(self.parts["xl/workbook.xml"])
        self.wb_rels = etree.fromstring(self.parts["xl/_rels/workbook.xml.rels"])
        self.content_types = etree.fromstring(self.parts["[Content_Types].xml"])
        self.styles = StyleBook(self.parts["xl/styles.xml"])
        self._load_shared_strings()
        self.sheets: Dict[str, Sheet] = {}
        self.sheet_paths: Dict[str, str] = {}
        rels = {r.get("Id"): r.get("Target") for r in self.wb_rels.findall("rel:Relationship", NS)}
        for s in self.workbook.find("m:sheets", NS).findall("m:sheet", NS):
            target = rels[s.get(R_ID)]
            p = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
            self.sheet_paths[s.get("name")] = p
        self._charts_cache: Dict[str, etree._Element] = {}
        self._drawings_cache: Dict[str, etree._Element] = {}
        self._dirty_parts: set = set()

    # -- shared strings ----------------------------------------------------------
    def _load_shared_strings(self):
        self.shared_strings: List[str] = []
        self._ss_index: Dict[str, int] = {}
        self._ss_root = None
        if "xl/sharedStrings.xml" in self.parts:
            self._ss_root = etree.fromstring(self.parts["xl/sharedStrings.xml"])
            for i, si in enumerate(self._ss_root.findall("m:si", NS)):
                txt = "".join(t.text or "" for t in si.iter(M + "t"))
                self.shared_strings.append(txt)
                self._ss_index.setdefault(txt, i)

    def add_shared_string(self, text: str) -> int:
        if text in self._ss_index:
            return self._ss_index[text]
        if self._ss_root is None:
            self._ss_root = etree.Element(M + "sst", nsmap={None: NS["m"]})
            self.parts["xl/sharedStrings.xml"] = b""
        si = etree.SubElement(self._ss_root, M + "si")
        t = etree.SubElement(si, M + "t")
        t.text = text
        if text != text.strip() or "\n" in text:
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        idx = len(self.shared_strings)
        self.shared_strings.append(text)
        self._ss_index[text] = idx
        return idx

    # -- hojas ---------------------------------------------------------------------
    def sheet_names(self) -> List[str]:
        return [s.get("name") for s in self.workbook.find("m:sheets", NS).findall("m:sheet", NS)]

    def sheet(self, name: str) -> Sheet:
        if name not in self.sheets:
            p = self.sheet_paths[name]
            self.sheets[name] = Sheet(self, name, p, self.parts[p])
        return self.sheets[name]

    def _rels_path(self, part: str) -> str:
        d, f = posixpath.split(part)
        return posixpath.join(d, "_rels", f + ".rels")

    def _rels(self, part: str) -> List[Tuple[str, str, str, Optional[str]]]:
        rp = self._rels_path(part)
        if rp not in self.parts:
            return []
        root = etree.fromstring(self.parts[rp])
        out = []
        for r in root.findall("rel:Relationship", NS):
            tgt = r.get("Target")
            mode = r.get("TargetMode")
            if mode != "External":
                tgt = posixpath.normpath(posixpath.join(posixpath.dirname(part), tgt))
            out.append((r.get("Id"), r.get("Type"), tgt, mode))
        return out

    def drawing_of(self, sheet_name: str) -> Optional[str]:
        for _, typ, tgt, _ in self._rels(self.sheet_paths[sheet_name]):
            if typ.endswith("/drawing"):
                return tgt
        return None

    def charts_of(self, sheet_name: str) -> List[str]:
        d = self.drawing_of(sheet_name)
        if not d:
            return []
        return [tgt for _, typ, tgt, _ in self._rels(d) if typ.endswith("/chart")]

    def all_chart_parts(self) -> List[str]:
        return [p for p in self.parts if re.match(r"xl/charts/chart\d+\.xml$", p)]

    def chart(self, part: str):
        if part not in self._charts_cache:
            self._charts_cache[part] = etree.fromstring(self.parts[part])
        self._dirty_parts.add(part)
        return self._charts_cache[part]

    def drawing(self, part: str):
        if part not in self._drawings_cache:
            self._drawings_cache[part] = etree.fromstring(self.parts[part])
        self._dirty_parts.add(part)
        return self._drawings_cache[part]

    # -- nombres definidos -------------------------------------------------------------
    def defined_names(self) -> Dict[str, str]:
        dn = self.workbook.find("m:definedNames", NS)
        if dn is None:
            return {}
        return {d.get("name"): d.text for d in dn.findall("m:definedName", NS) if d.get("localSheetId") is None}

    def set_defined_name(self, name: str, ref: str, comment: Optional[str] = None):
        dn = self.workbook.find("m:definedNames", NS)
        if dn is None:
            dn = etree.Element(M + "definedNames")
            self.workbook.find("m:sheets", NS).addnext(dn)
        for d in dn.findall("m:definedName", NS):
            if d.get("name") == name and d.get("localSheetId") is None:
                d.text = ref
                if comment:
                    d.set("comment", comment)
                return
        el = etree.Element(M + "definedName", name=name)
        if comment:
            el.set("comment", comment)
        el.text = ref
        # Excel exige orden alfabético (no estricto, pero lo respetamos)
        placed = False
        for d in dn.findall("m:definedName", NS):
            if d.get("name").lower() > name.lower():
                d.addprevious(el)
                placed = True
                break
        if not placed:
            dn.append(el)

    # -- propagación de inserción de filas ------------------------------------------------
    def _after_rows_inserted(self, sheet: Sheet, at_row: int, n: int):
        name = sheet.name
        # fórmulas en todas las hojas
        for sname in self.sheet_names():
            sh = self.sheet(sname)
            for f in sh.root.iter(M + "f"):
                if f.text:
                    f.text = shift_formula_rows(f.text, name, sname, at_row, n)
                if f.get("ref"):
                    if sname == name:
                        f.set("ref", shift_sqref(f.get("ref"), at_row, n))
        # nombres definidos
        dn = self.workbook.find("m:definedNames", NS)
        if dn is not None:
            for d in dn.findall("m:definedName", NS):
                if d.text and "!" in d.text:
                    d.text = shift_formula_rows(d.text, name, "\0", at_row, n)
        # gráficos
        for part in self.all_chart_parts():
            ch = self.chart(part)
            for f in ch.iter("{%s}f" % NS["c"]):
                if f.text:
                    f.text = shift_formula_rows(f.text, name, "\0", at_row, n)
        # anclas de dibujos de la hoja
        d = self.drawing_of(name)
        if d:
            dr = self.drawing(d)
            for tag in ("from", "to"):
                for el in dr.iter("{%s}%s" % (NS["xdr"], tag)):
                    row = el.find("xdr:row", NS)
                    if row is not None and int(row.text) >= at_row - 1:
                        row.text = str(int(row.text) + n)
        # calcChain queda obsoleto
        self.drop_calc_chain()

    def drop_calc_chain(self):
        if "xl/calcChain.xml" in self.parts:
            del self.parts["xl/calcChain.xml"]
            self.order = [o for o in self.order if o != "xl/calcChain.xml"]
            for r in self.wb_rels.findall("rel:Relationship", NS):
                if r.get("Target", "").endswith("calcChain.xml"):
                    self.wb_rels.remove(r)
            for o in self.content_types.findall("ct:Override", NS):
                if o.get("PartName") == "/xl/calcChain.xml":
                    self.content_types.remove(o)

    def set_active_sheet(self, name: str):
        """Deja una sola hoja seleccionada (si activeTab y tabSelected no coinciden, Excel abre hojas agrupadas)."""
        names = self.sheet_names()
        if name not in names:
            return
        bv = self.workbook.find("m:bookViews/m:workbookView", NS)
        if bv is not None:
            bv.set("activeTab", str(names.index(name)))
            bv.set("firstSheet", "0")
        for sname in names:
            sh = self.sheet(sname)
            for sv in sh.root.findall("m:sheetViews/m:sheetView", NS):
                if sname == name:
                    sv.set("tabSelected", "1")
                elif "tabSelected" in sv.attrib:
                    del sv.attrib["tabSelected"]

    def set_full_calc_on_load(self):
        cp = self.workbook.find("m:calcPr", NS)
        if cp is None:
            cp = etree.Element(M + "calcPr")
            # calcPr va después de definedNames (o de sheets si no hay nombres)
            anchor = self.workbook.find("m:definedNames", NS)
            if anchor is None:
                anchor = self.workbook.find("m:sheets", NS)
            anchor.addnext(cp)
        cp.set("fullCalcOnLoad", "1")

    # -- hojas nuevas -----------------------------------------------------------------------
    def add_or_replace_sheet(self, name: str, after: Optional[str] = None) -> Sheet:
        """Crea una hoja vacía (o la vacía si existe) y la devuelve."""
        empty = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="%s" xmlns:r="%s"><sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="14.5"/><sheetData/>'
            '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/></worksheet>'
            % (NS["m"], NS["r"])
        ).encode("utf-8")
        if name in self.sheet_paths:
            p = self.sheet_paths[name]
            self.parts[p] = empty
            rp = self._rels_path(p)
            if rp in self.parts:
                del self.parts[rp]
                self.order = [o for o in self.order if o != rp]
            self.sheets.pop(name, None)
            return self.sheet(name)
        # nueva parte
        nums = [int(m.group(1)) for p in self.parts for m in [re.match(r"xl/worksheets/sheet(\d+)\.xml$", p)] if m]
        p = f"xl/worksheets/sheet{max(nums) + 1}.xml"
        self.parts[p] = empty
        self.order.append(p)
        rid_nums = [int(r.get("Id")[3:]) for r in self.wb_rels.findall("rel:Relationship", NS) if r.get("Id", "").startswith("rId")]
        rid = f"rId{max(rid_nums) + 1}"
        etree.SubElement(self.wb_rels, "{%s}Relationship" % NS["rel"], Id=rid, Type=REL_WORKSHEET,
                         Target=p.replace("xl/", "", 1))
        etree.SubElement(self.content_types, "{%s}Override" % NS["ct"], PartName="/" + p, ContentType=CT_WORKSHEET)
        sheets = self.workbook.find("m:sheets", NS)
        ids = [int(s.get("sheetId")) for s in sheets.findall("m:sheet", NS)]
        el = etree.Element(M + "sheet", name=name, sheetId=str(max(ids) + 1))
        el.set(R_ID, rid)
        if after and after in self.sheet_paths:
            for s in sheets.findall("m:sheet", NS):
                if s.get("name") == after:
                    s.addnext(el)
                    break
        else:
            sheets.append(el)
        self.sheet_paths[name] = p
        self._update_app_titles()
        return self.sheet(name)

    def _update_app_titles(self):
        """Mantiene docProps/app.xml coherente con la lista de hojas (Excel lo tolera, pero mejor así)."""
        p = "docProps/app.xml"
        if p not in self.parts:
            return
        try:
            root = etree.fromstring(self.parts[p])
            names = self.sheet_names()
            hp = root.find("ep:HeadingPairs", NS)
            tp = root.find("ep:TitlesOfParts", NS)
            if hp is None or tp is None:
                return
            variants = hp.find("vt:vector", NS).findall("vt:variant", NS)
            n_sheets_old = None
            for i in range(0, len(variants) - 1, 2):
                lp = variants[i].find("vt:lpstr", NS)
                if lp is not None and lp.text and ("Hoja" in lp.text or "Worksheet" in lp.text):
                    i4 = variants[i + 1].find("vt:i4", NS)
                    n_sheets_old = int(i4.text)
                    i4.text = str(len(names))
            if n_sheets_old is None:
                return
            tvec = tp.find("vt:vector", NS)
            existing = [e.text for e in tvec.findall("vt:lpstr", NS)]
            # en TitlesOfParts primero van las hojas y luego otros elementos (rangos con nombre)
            tail = existing[n_sheets_old:]
            for e in list(tvec):
                tvec.remove(e)
            for nm in names + tail:
                etree.SubElement(tvec, "{%s}lpstr" % NS["vt"]).text = nm
            tvec.set("size", str(len(names) + len(tail)))
            self.parts[p] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        except Exception:
            pass

    def add_hyperlink(self, sheet: Sheet, ref: str, url: str):
        """Hipervínculo externo en una celda."""
        rp = self._rels_path(sheet.path)
        if rp in self.parts:
            root = etree.fromstring(self.parts[rp])
        else:
            root = etree.Element("{%s}Relationships" % NS["rel"], nsmap={None: NS["rel"]})
            self.order.append(rp)
        nums = [int(r.get("Id")[3:]) for r in root.findall("rel:Relationship", NS) if r.get("Id", "").startswith("rId")]
        rid = f"rId{(max(nums) if nums else 0) + 1}"
        etree.SubElement(root, "{%s}Relationship" % NS["rel"], Id=rid, Type=REL_HYPERLINK, Target=url,
                         TargetMode="External")
        self.parts[rp] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        hls = sheet.root.find("m:hyperlinks", NS)
        if hls is None:
            hls = etree.Element(M + "hyperlinks")
            # orden del esquema: ... dataValidations, hyperlinks, printOptions, pageMargins ...
            nxt = None
            for tag in ("printOptions", "pageMargins", "pageSetup", "headerFooter", "drawing"):
                nxt = sheet.root.find("m:" + tag, NS)
                if nxt is not None:
                    break
            if nxt is not None:
                nxt.addprevious(hls)
            else:
                sheet.root.append(hls)
        h = etree.SubElement(hls, M + "hyperlink", ref=ref)
        h.set(R_ID, rid)

    # -- guardar ---------------------------------------------------------------------------------
    def save(self, out_path: str):
        for name, sh in self.sheets.items():
            self.parts[sh.path] = sh.tobytes()
        for part in self._dirty_parts:
            if part in self._charts_cache:
                self.parts[part] = etree.tostring(self._charts_cache[part], xml_declaration=True,
                                                  encoding="UTF-8", standalone=True)
            elif part in self._drawings_cache:
                self.parts[part] = etree.tostring(self._drawings_cache[part], xml_declaration=True,
                                                  encoding="UTF-8", standalone=True)
        if self._ss_root is not None:
            n = len(self._ss_root.findall("m:si", NS))
            self._ss_root.set("uniqueCount", str(n))
            self._ss_root.set("count", str(max(n, int(self._ss_root.get("count", "0") or 0))))
            self.parts["xl/sharedStrings.xml"] = etree.tostring(self._ss_root, xml_declaration=True,
                                                                encoding="UTF-8", standalone=True)
        self.parts["xl/styles.xml"] = self.styles.tobytes()
        self.parts["xl/workbook.xml"] = etree.tostring(self.workbook, xml_declaration=True, encoding="UTF-8",
                                                       standalone=True)
        self.parts["xl/_rels/workbook.xml.rels"] = etree.tostring(self.wb_rels, xml_declaration=True,
                                                                  encoding="UTF-8", standalone=True)
        self.parts["[Content_Types].xml"] = etree.tostring(self.content_types, xml_declaration=True,
                                                           encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            written = set()
            for n in self.order:
                if n in self.parts and n not in written:
                    z.writestr(n, self.parts[n])
                    written.add(n)
            for n, data in self.parts.items():
                if n not in written:
                    z.writestr(n, data)
