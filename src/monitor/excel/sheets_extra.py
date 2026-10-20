"""Hojas generadas por el pipeline: Noticias y Log_Actualizacion, y ajustes en Resumen."""
from __future__ import annotations

import datetime as dt
from typing import Dict, List, Optional

from lxml import etree

from .xlsx_book import M, NS, Sheet, XlsxBook, idx_to_col

NAVY = "FF002060"
GREY = "FFF2F2F2"
LIGHT = "FFDDEBF7"


class Styles:
    def __init__(self, book: XlsxBook):
        sb = book.styles
        self.title = sb.new_style(bold=True, size=14, color=NAVY)
        self.subtitle = sb.new_style(italic=True, size=9, color="FF595959")
        self.h_country = sb.new_style(bold=True, size=11, color="FFFFFFFF", fill_rgb=NAVY, valign="center")
        self.header = sb.new_style(bold=True, size=9, color="FF000000", fill_rgb=LIGHT, wrap=True, valign="center",
                                   border=True)
        self.text = sb.new_style(size=9, wrap=True, border=True)
        self.text_nb = sb.new_style(size=9, wrap=True)
        self.text_bold = sb.new_style(size=9, bold=True, wrap=True, border=True)
        self.link = sb.new_style(size=9, color="FF0563C1", underline=True, wrap=True, border=True)
        self.date = sb.new_style(size=9, num_fmt="dd-mmm-yy", border=True, halign="center")
        self.pct = sb.new_style(size=9, num_fmt="0.00%;[Red]-0.00%", border=True, halign="center")
        self.num = sb.new_style(size=9, num_fmt="0.00;[Red]-0.00", border=True, halign="center")
        self.pp = sb.new_style(size=9, num_fmt='+0.00" pp";[Red]-0.00" pp"', border=True, halign="center")
        self.pts = sb.new_style(size=9, num_fmt='+0.00" pts";[Red]-0.00" pts"', border=True, halign="center")
        self.ok = sb.new_style(size=9, bold=True, color="FF008000", border=True, halign="center")
        self.warn = sb.new_style(size=9, bold=True, color="FFC55A11", border=True, halign="center")
        self.err = sb.new_style(size=9, bold=True, color="FFC00000", border=True, halign="center")
        self.narr = sb.new_style(size=10, wrap=True, valign="top", fill_rgb=GREY)


def _set_cols(sheet: Sheet, widths: Dict[str, float]):
    cols = sheet.root.find("m:cols", NS)
    if cols is None:
        cols = etree.Element(M + "cols")
        sheet.root.find("m:sheetData", NS).addprevious(cols)
    for c in list(cols):
        cols.remove(c)
    from .xlsx_book import col_to_idx
    for col, w in widths.items():
        i = col_to_idx(col)
        etree.SubElement(cols, M + "col", min=str(i), max=str(i), width=str(w), customWidth="1")


def _row_height(sheet: Sheet, r: int, h: float):
    row = sheet.row_el(r, create=True)
    row.set("ht", str(h))
    row.set("customHeight", "1")


def _merge(sheet: Sheet, ref: str):
    mc = sheet.root.find("m:mergeCells", NS)
    if mc is None:
        mc = etree.Element(M + "mergeCells")
        sheet.root.find("m:sheetData", NS).addnext(mc)
    etree.SubElement(mc, M + "mergeCell", ref=ref)
    mc.set("count", str(len(mc)))


def _freeze(sheet: Sheet, row: int):
    sv = sheet.root.find("m:sheetViews/m:sheetView", NS)
    pane = etree.SubElement(sv, M + "pane", ySplit=str(row - 1), topLeftCell=f"A{row}", activePane="bottomLeft",
                            state="frozen")
    etree.SubElement(sv, M + "selection", pane="bottomLeft", activeCell=f"A{row}", sqref=f"A{row}")


# ------------------------------------------------------------------------------------- Log
def build_log_sheet(book: XlsxBook, rows: List[dict], changes: list, run_info: dict):
    sh = book.add_or_replace_sheet("Log_Actualizacion", after="7_TPM's")
    st = Styles(book)
    _set_cols(sh, {"A": 2, "B": 24, "C": 16, "D": 13, "E": 30, "F": 12, "G": 11, "H": 30, "I": 60})
    sh.set_text("B2", "Log de actualización", st.title)
    sh.set_text("B3", f"Corrida: {run_info['inicio']}  |  modo: {run_info['modo']}  |  "
                      f"series OK: {run_info['ok']} de {run_info['total']}  |  celdas nuevas: {run_info['nuevos']}  |  "
                      f"revisiones: {run_info['revisiones']}", st.subtitle)
    hdr = ["Hoja", "Serie", "Estado", "Fuente usada", "Último dato", "Meses escritos", "Validación", "Detalle / intentos"]
    r = 5
    for i, h in enumerate(hdr):
        sh.set_text(f"{idx_to_col(i + 2)}{r}", h, st.header)
    _row_height(sh, r, 24)
    for row in rows:
        r += 1
        estado = row["estado"]
        est_style = st.ok if estado == "OK" else (st.warn if estado in ("MANUAL", "NOTICIAS", "SIN CAMBIOS") else st.err)
        vals = [row["hoja"], row["serie"], estado, row.get("fuente", ""), row.get("ultimo", ""),
                row.get("escritos", ""), row.get("validacion", ""), row.get("detalle", "")]
        for i, v in enumerate(vals):
            ref = f"{idx_to_col(i + 2)}{r}"
            if i == 5 and isinstance(v, (int, float)):
                sh.set_number(ref, v, st.text)
            else:
                sh.set_text(ref, "" if v is None else str(v), est_style if i == 2 else st.text)
    _freeze(sh, 6)
    # cambios
    r += 3
    sh.set_text(f"B{r}", "Celdas modificadas en esta corrida (nuevos datos y revisiones de la fuente)", st.title)
    r += 1
    for i, h in enumerate(["Hoja", "Serie", "Mes", "Celda", "Anterior", "Nuevo", "Tipo", "Fuente"]):
        sh.set_text(f"{idx_to_col(i + 2)}{r}", h, st.header)
    for ch in changes[:3000]:
        r += 1
        sh.set_text(f"B{r}", ch.sheet, st.text)
        sh.set_text(f"C{r}", ch.column, st.text)
        sh.set_text(f"D{r}", ch.month, st.text)
        sh.set_text(f"E{r}", ch.cell, st.text)
        sh.set_number(f"F{r}", ch.old, st.num if (ch.new or 0) > 1.5 else st.pct)
        sh.set_number(f"G{r}", ch.new, st.num if (ch.new or 0) > 1.5 else st.pct)
        sh.set_text(f"H{r}", ch.kind, st.text)
        sh.set_text(f"I{r}", ch.source, st.text)
    return sh


# ------------------------------------------------------------------------------------- Noticias
def build_news_sheet(book: XlsxBook, blocks: List[dict], run_info: dict):
    """blocks: [{pais, narrativa, indicadores:[{...}], noticias:[NewsItem]}]"""
    sh = book.add_or_replace_sheet("Noticias", after="Resumen")
    st = Styles(book)
    _set_cols(sh, {"A": 2, "B": 22, "C": 11, "D": 12, "E": 12, "F": 12, "G": 12, "H": 62, "I": 18, "J": 46})
    sh.set_text("B2", "Principales noticias y lectura por país", st.title)
    sh.set_text("B3", f"Actualizado: {run_info['inicio']}. Titulares de los últimos {run_info.get('dias', 45)} días "
                      "(Google News RSS + feeds oficiales). La columna «Por qué» es una extracción automática del "
                      "titular; verificar antes de citar.", st.subtitle)
    r = 5
    for b in blocks:
        _merge(sh, f"B{r}:J{r}")
        sh.set_text(f"B{r}", b["pais"], st.h_country)
        for col in "CDEFGHIJ":
            sh.set_text(f"{col}{r}", None, st.h_country)
        _row_height(sh, r, 18)
        r += 1
        if b.get("narrativa"):
            _merge(sh, f"B{r}:J{r}")
            sh.set_text(f"B{r}", b["narrativa"], st.narr)
            for col in "CDEFGHIJ":
                sh.set_text(f"{col}{r}", None, st.narr)
            lines = max(2, len(b["narrativa"]) // 150 + 1)
            _row_height(sh, r, 14 * lines)
            r += 1
        # tabla de datos
        hdr = ["Indicador", "Último mes", "Dato", "Mes previo", "Cambio m/m", "Hace 12m", "Lectura"]
        for i, h in enumerate(hdr):
            sh.set_text(f"{idx_to_col(i + 2)}{r}", h, st.header)
        r += 1
        for ind in b["indicadores"]:
            is_pct = ind["unidad"] == "pct"
            vs = st.pct if is_pct else st.num
            ds = st.pp if is_pct else st.pts
            sh.set_text(f"B{r}", ind["nombre"], st.text_bold)
            sh.set_text(f"C{r}", ind["mes"], st.text)
            sh.set_number(f"D{r}", ind["valor"], vs)
            sh.set_number(f"E{r}", ind.get("previo"), vs)
            delta = ind.get("delta")
            sh.set_number(f"F{r}", None if delta is None else (delta * 100 if is_pct else delta), ds)
            sh.set_number(f"G{r}", ind.get("hace12"), vs)
            sh.set_text(f"H{r}", ind.get("lectura", ""), st.text)
            r += 1
        # noticias
        if b["noticias"]:
            r += 0
            for i, h in enumerate(["Indicador", "Fecha", "Fuente", "", "", "", "Titular", "Por qué (extracto)", ""]):
                if h:
                    sh.set_text(f"{idx_to_col(i + 2)}{r}", h, st.header)
            _merge(sh, f"E{r}:G{r}")
            sh.set_text(f"E{r}", None, st.header)
            r += 1
            for it in b["noticias"]:
                sh.set_text(f"B{r}", it.tags[0] if it.tags else "", st.text)
                if it.published:
                    sh.set_date(f"C{r}", it.published.date(), st.date)
                else:
                    sh.set_text(f"C{r}", "", st.text)
                _merge(sh, f"D{r}:G{r}")
                sh.set_text(f"D{r}", it.source, st.text)
                for col in "EFG":
                    sh.set_text(f"{col}{r}", None, st.text)
                sh.set_text(f"H{r}", it.title, st.link)
                if it.link:
                    book.add_hyperlink(sh, f"H{r}", it.link)
                _merge(sh, f"I{r}:J{r}")
                sh.set_text(f"I{r}", it.why, st.text)
                sh.set_text(f"J{r}", None, st.text)
                _row_height(sh, r, 26 if len(it.title) > 70 else 14.5)
                r += 1
        r += 1
    return sh


# ------------------------------------------------------------------------------------- Resumen
STATUS_MAP = {"OK": ("OFICIAL", "FF008000"), "MANUAL": ("MANUAL", "FFC55A11"), "NOTICIAS": ("PRENSA", "FFC55A11"),
              "ERROR": ("REVISAR", "FFC00000"), "SIN FUENTE": ("MANUAL", "FFC55A11")}


def update_resumen(book: XlsxBook, run_info: dict, status: Dict[str, Dict[str, str]]):
    if "Resumen" not in book.sheet_names():
        return
    sh = book.sheet("Resumen")
    sh.set_text("B4", f"Última actualización automática: {run_info['inicio']} (python -m monitor)", sh.style("B4"))
    steps = [
        "1. Doble clic en run.bat (o en VS Code: Terminal > Run Task > Actualizar Monitor).",
        "2. El código descarga cada serie de su fuente oficial, la valida contra el histórico y la escribe aquí.",
        "3. Heatmaps (ventana móvil), promedios y gráficos se ajustan solos al último mes con dato.",
        "4. Revisa la hoja Log_Actualizacion: estado por serie, fuente usada y celdas que cambiaron.",
        "5. Si una fuente falla, escribe el dato en data/manual_inputs.csv y vuelve a correr.",
        "6. Las celdas en AMARILLO son meses aún no publicados; texto NARANJA = dato manual/noticias por verificar.",
    ]
    for i, txt in enumerate(steps):
        ref = f"B{17 + i}"
        sh.set_text(ref, txt, sh.style(ref))
    if isinstance(sh.get("B9"), str) and "VERDE" in sh.get("B9").upper():
        sh.set_text("C9", "Dato oficial: descargado automáticamente de la fuente por el código (o verificado en ella).",
                    sh.style("C9"))
    # leyenda adicional (fila 14 libre en la plantilla)
    if sh.get("B14") in (None, ""):
        base = sh.style("B12")
        st_or = book.styles.derive(base, font_rgb="FFC55A11") if base is not None else None
        sh.set_text("B14", "Texto NARANJA", st_or)
        sh.set_text("C14", "Dato manual sin verificar o extraído de titulares de prensa: confirmar contra el comunicado oficial.",
                    sh.style("C13"))
    # tabla de estado
    hojas = {sh.get(f"B{r}"): r for r in range(32, 40) if isinstance(sh.get(f"B{r}"), str)}
    paises = {"BR": "BRASIL", "CL": "CHILE", "CO": "COLOMBIA", "MX": "MEXICO", "PE": "PERU"}
    cols = {"BR": "C", "CL": "D", "CO": "E", "MX": "F", "PE": "G"}
    for hoja, r in hojas.items():
        st_h = status.get(hoja, {})
        for code, col in cols.items():
            key = paises[code]
            if hoja == "6_ANTAD":
                key = "OVERALL" if code == "MX" else None
            if key and key in st_h:
                label, color = STATUS_MAP.get(st_h[key], ("REVISAR", "FFC00000"))
                base = sh.style(f"{col}{r}")
                style = book.styles.derive(base, font_rgb=color) if base is not None else None
                sh.set_text(f"{col}{r}", label, style)
