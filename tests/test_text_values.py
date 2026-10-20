"""Extracción de datos desde el texto de boletines y titulares."""
import logging
from pathlib import Path

import pandas as pd

from monitor.http import Http
from monitor.pipeline import Context
from monitor.sources import REGISTRY
from monitor.sources.text_values import extract_value, find_month


def test_frases_de_boletines():
    t = ("En septiembre de 2026 y con datos ajustados por estacionalidad, el Indicador de Confianza del "
         "Consumidor (ICC) se ubicó en 45.1 puntos, una disminución mensual de 1.2 puntos.")
    assert extract_value(t, "Indicador de Confianza del Consumidor", "puntos", (20, 80))[0] == 45.1
    t = ("En julio de 2026pr, el Indicador de Seguimiento a la Economía (ISE) en su serie original se ubicó en "
         "131,74 y presentó un crecimiento de 1,11 % respecto al mismo mes de 2025pr.")
    assert extract_value(t, "serie original", "%", (-30, 40))[0] == 1.11
    t = "Confianza comercial bajó a 23,3 e industrial a 5,8 en agosto según Fedesarrollo"
    assert extract_value(t, "confianza comercial", "auto", (-60, 80), "cercano")[0] == 23.3
    t = "En agosto, el ICC de Fedesarrollo cayó a -5,2% frente a julio"
    assert extract_value(t, "ICC", "auto", (-60, 60), "cercano")[0] == -5.2
    assert find_month("O ICEI subiu para 46,2 pontos em outubro de 2026") == (2026, 10)


def test_texto_web_html(tmp_path):
    latest = tmp_path / "latest"
    latest.mkdir()
    (latest / "cni_test_icei.html").write_text(
        "<html><body><h1>ICEI</h1><p>O ICEI ficou em 46,2 pontos em outubro de 2026, alta de 1,3 ponto.</p>"
        "<script>var x = 99.9;</script></body></html>", encoding="utf-8")
    ctx = Context(Http(tmp_path, offline=True, use_system_certs=False), logging.getLogger("t"))
    r = REGISTRY["texto_web"]({"clave": "cni", "formato": "html", "url": "https://x/test_icei", "patron": "ICEI",
                               "unidad": "puntos", "rango": [20, 80]}, ctx)
    assert r.series[pd.Timestamp(2026, 10, 1)] == 46.2
    assert r.verified is False
