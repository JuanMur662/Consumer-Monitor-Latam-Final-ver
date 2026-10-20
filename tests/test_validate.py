import pandas as pd

from monitor.sources.base import apply_transform, parse_period
from monitor.validate import validate


def test_transforms():
    idx = pd.date_range("2024-01-01", periods=24, freq="MS")
    s = pd.Series(range(100, 124), index=idx, dtype=float)
    y = apply_transform(s, "yoy")
    assert abs(y.iloc[0] - (112 / 100 - 1)) < 1e-12
    assert apply_transform(pd.Series([20.0]), "formula:(x+100)/2").iloc[0] == 60
    assert apply_transform(pd.Series([127.39]), "formula:x/2").iloc[0] == 63.695
    assert apply_transform(pd.Series([4.5]), "pct").iloc[0] == 0.045


def test_periodos():
    assert parse_period("Ene.2026") == pd.Timestamp(2026, 1, 1)
    assert parse_period("Set.2025") == pd.Timestamp(2025, 9, 1)
    assert parse_period("2026/08") == pd.Timestamp(2026, 8, 1)
    assert parse_period("01/08/2026") == pd.Timestamp(2026, 8, 1)
    assert parse_period("agosto 2026") == pd.Timestamp(2026, 8, 1)
    assert parse_period("202608") == pd.Timestamp(2026, 8, 1)
    assert parse_period("2018.5") is None


def test_validacion_rechaza_serie_equivocada():
    idx = pd.date_range("2024-01-01", periods=24, freq="MS")
    hist = pd.Series([0.03 + 0.001 * i for i in range(24)], index=idx)
    ok = validate(hist + 0.0005, hist, mad_max=0.003)
    assert ok.ok
    mala = validate(hist * 100, hist, mad_max=0.003, rango=(-0.05, 0.4))   # error de escala
    assert not mala.ok
    otra = validate(hist + 0.02, hist, mad_max=0.003)                        # otra serie / definición
    assert not otra.ok
