"""El Excel actualizado conserva gráficos, formatos y fórmulas; solo cambian datos y rangos."""
import shutil
import zipfile
from pathlib import Path

import pandas as pd

from monitor.excel.monitor_sheets import DataTable, Heatmap, update_charts
from monitor.excel.xlsx_book import XlsxBook

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "excel" / "Monitor_Retail_LATAM_original.xlsx"


def test_escribe_y_extiende(tmp_path):
    src = tmp_path / "in.xlsx"
    shutil.copy(TEMPLATE, src)
    b = XlsxBook(str(src))
    t = DataTable.locate(b, "1_Actividad_Economica")
    n_rows_before = len(t.rows)
    s = pd.Series([0.011, 0.022], index=pd.to_datetime(["2026-08-01", "2026-12-01"]))
    ch = t.write_series("BRASIL", s, "test")
    assert {c.month for c in ch} == {"2026-08", "2026-12"}
    assert len(t.rows) == n_rows_before + 7          # filas jun..dic 2026 insertadas
    hm = Heatmap.locate(t.sheet)
    assert hm.refresh(t).month == 12
    assert t.sheet.get("B18") == "Dic-26"
    assert update_charts(b, t) == 10
    out = tmp_path / "out.xlsx"
    b.save(str(out))
    # notas desplazadas, merges desplazados, gráfico extendido
    b2 = XlsxBook(str(out))
    sh = b2.sheet("1_Actividad_Economica")
    assert sh.get("B95") == "Fuentes:"
    chart = b2.parts["xl/charts/chart1.xml"].decode()
    assert "$B$24:$B$93" in chart
    with zipfile.ZipFile(out) as z:
        names = set(z.namelist())
    with zipfile.ZipFile(TEMPLATE) as z:
        orig = set(z.namelist())
    assert {n for n in orig if "chart" in n or "drawing" in n} <= names   # nada de gráficos se pierde


def test_estilos_y_formulas_compartidas(tmp_path):
    src = tmp_path / "in.xlsx"
    shutil.copy(TEMPLATE, src)
    b = XlsxBook(str(src))
    t = DataTable.locate(b, "5_Inflacion")
    t.normalize_formula_dates()
    assert all(d.day == 1 for _, d in t.rows)      # sin =EOMONTH mezclado
    hm = Heatmap.locate(t.sheet)
    hm.refresh(t)
    # no deben quedar fórmulas compartidas huérfanas
    import re
    xml = t.sheet.tobytes().decode()
    masters = set(re.findall(r'<f t="shared" ref="[^"]+" si="(\d+)"', xml))
    children = set(re.findall(r'<f t="shared" si="(\d+)"\s*/>', xml))
    assert children <= masters
