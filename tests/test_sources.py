"""Parsers de las APIs con respuestas reales guardadas (tests/fixtures/latest)."""
import logging
from pathlib import Path

import pandas as pd

from monitor.http import Http
from monitor.pipeline import Context
from monitor.sources import REGISTRY, apply_transform

FIX = Path(__file__).resolve().parent / "fixtures"


def ctx():
    return Context(Http(FIX, offline=True, use_system_certs=False), logging.getLogger("t"))


def test_bcb_ibcbr_yoy():
    r = REGISTRY["bcb_sgs"]({"codigo": 24363}, ctx())
    y = apply_transform(r.series, "yoy")
    assert abs(y[pd.Timestamp(2026, 3, 1)] - (118.01719 / 113.99468 - 1)) < 1e-9


def test_bcrp():
    r = REGISTRY["bcrp"]({"codigo": "PN01273PM"}, ctx())
    assert abs(r.series[pd.Timestamp(2026, 9, 1)] - 4.54865469481509) < 1e-9


def test_bis_y_fred():
    r = REGISTRY["bis"]({"dataset": "WS_CBPOL", "clave": "M.CO"}, ctx())
    assert r.series[pd.Timestamp(2026, 7, 1)] == 12
    f = REGISTRY["fred"]({"codigo": "FEDFUNDS"}, ctx())
    assert f.series[pd.Timestamp(2026, 9, 1)] == 3.75


def test_sidra():
    r = REGISTRY["ibge_sidra"]({"tabela": 8880, "variavel": 11709, "classificacao": "c11046/56734"}, ctx())
    assert r.series[pd.Timestamp(2025, 4, 1)] == 5.4


def test_bcch_siete():
    r = REGISTRY["bcch_siete"]({"cuadro": "CAP_BDP/MN_EXP_EC11/EXE_IMCE_TOTAL", "fila": "IMCE: ?Comercio"}, ctx())
    assert r.series[pd.Timestamp(2026, 2, 1)] == 57.67
    r = REGISTRY["bcch_siete"]({"cuadro": "CAP_CCNN/MN_CCNN76/CCNN2018_IMACEC_01_A", "calculo": "YTYPCT",
                                "fila": "^Imacec$"}, ctx())
    assert r.series[pd.Timestamp(2026, 8, 1)] == -1.0


def test_noticia_dato_antad():
    from monitor.news import parse_rss
    xml = (b'<rss><channel><item><title>Ventas ANTAD crecen 4.1% a tiendas totales y caen 1.2% a tiendas iguales '
           b'en septiembre - El Economista</title><source>El Economista</source>'
           b'<pubDate>Fri, 09 Oct 2026 18:00:00 GMT</pubDate><link>https://x</link></item></channel></rss>')
    items = parse_rss(xml)
    assert items[0].source == "El Economista"


def test_noticia_dato_valor_y_signo():
    """Toma el número más cercano a 'tiendas iguales', detecta 'caen/retroceden' y exige el mes más reciente."""
    r = REGISTRY["noticia_dato"]({"consulta": "ANTAD ventas tiendas iguales", "pais": "MX",
                                  "patron": "tiendas iguales", "rango": [-30, 40]}, ctx())
    assert r.series.index[-1] == pd.Timestamp(2026, 9, 1)
    assert r.series.iloc[-1] == -1.2
    assert r.verified is False and "2 fuentes" in r.source
