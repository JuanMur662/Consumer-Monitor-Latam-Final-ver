"""El lector genérico debe encontrar la serie en las disposiciones típicas de DANE / INE / CNI."""
import io

import pandas as pd
from openpyxl import Workbook

from monitor.sources.tables import find_series, read_all_sheets


def _xlsx(rows, title="Cuadro 1"):
    wb = Workbook()
    ws = wb.active
    ws.title = title
    for r in rows:
        ws.append(r)
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


def test_ancha_dos_filas_anio_mes():
    years = ["Concepto"] + [2025] * 12 + [2026] * 3
    months = [""] + ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre",
                     "Octubre", "Noviembre", "Diciembre", "Enero", "Febrero", "Marzo"]
    rows = [["Indicador de Seguimiento a la Economía (ISE) - Series originales"], [], years, months,
            ["Indicador de Seguimiento a la Economía"] + list(range(100, 115)),
            ["Actividades primarias"] + list(range(50, 65))]
    sheets = read_all_sheets(_xlsx(rows), "xlsx")
    c = find_series(sheets, "indicador de seguimiento")
    assert c and c[0].layout == "ancha"
    s = c[0].series
    assert s.index.min() == pd.Timestamp(2025, 1, 1) and s.index.max() == pd.Timestamp(2026, 3, 1)
    assert s[pd.Timestamp(2026, 3, 1)] == 114


def test_larga_columna_fechas():
    rows = [["Serie histórica ICEI"], ["Mes", "ICEI", "Condições atuais"]]
    for i in range(18):
        rows.append([f"{['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'][i % 12]}/{25 + i // 12}",
                     45 + i * 0.1, 40])
    sheets = read_all_sheets(_xlsx(rows), "xlsx")
    c = find_series(sheets, "^icei$|icei")
    assert c and c[0].layout == "larga"
    assert abs(c[0].series.iloc[-1] - (45 + 17 * 0.1)) < 1e-9


def test_matriz_mes_por_anio():
    rows = [["Índice de Precios al Consumidor (IPC) - Índices total nacional"], ["Mes", 2023, 2024, 2025, 2026]]
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre",
             "Noviembre", "Diciembre"]
    for i, m in enumerate(meses):
        rows.append([m, 130 + i, 140 + i, 147 + i, 155 + i if i < 9 else None])
    sheets = read_all_sheets(_xlsx(rows, "IPC"), "xlsx")
    c = find_series(sheets, "indice|ipc")
    assert c
    s = c[0].series.dropna()
    assert s.index.max() == pd.Timestamp(2026, 9, 1)
    assert s[pd.Timestamp(2025, 12, 1)] == 158
