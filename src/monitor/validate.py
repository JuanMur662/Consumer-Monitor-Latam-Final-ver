"""
Validación de una serie descargada contra el histórico que ya está en el Excel.

Objetivo: que un id equivocado, un cambio de formato del archivo fuente o un
error de escala NUNCA termine escrito en el Excel. Se compara el tramo en que
ambas series tienen datos (últimos 36 meses) y se exige:
  - diferencia mediana absoluta (MAD) <= mad_max
  - correlación >= corr_min (si hay al menos 8 meses de solape)
  - valores dentro de un rango plausible
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

import numpy as np
import pandas as pd


@dataclass
class Validation:
    ok: bool
    n_overlap: int
    mad: Optional[float]
    corr: Optional[float]
    reason: str

    def short(self) -> str:
        parts = [f"solape={self.n_overlap}"]
        if self.mad is not None:
            parts.append(f"MAD={self.mad:.4g}")
        if self.corr is not None:
            parts.append(f"corr={self.corr:.2f}")
        return ", ".join(parts) + (f" | {self.reason}" if self.reason else "")


def validate(new: pd.Series, old: pd.Series, mad_max: float, corr_min: float = 0.6,
             rango=(None, None), min_overlap: int = 3, window: int = 36) -> Validation:
    new = new.dropna()
    if new.empty:
        return Validation(False, 0, None, None, "serie vacía")
    lo, hi = rango
    if lo is not None and (new < lo).any():
        return Validation(False, 0, None, None, f"valores bajo el rango plausible ({new.min():.4g} < {lo})")
    if hi is not None and (new > hi).any():
        return Validation(False, 0, None, None, f"valores sobre el rango plausible ({new.max():.4g} > {hi})")
    old = old.dropna()
    common = new.index.intersection(old.index)
    common = common.sort_values()[-window:]
    n = len(common)
    if n < min_overlap:
        return Validation(True, n, None, None, "sin solape suficiente: aceptada solo por rango plausible")
    diff = (new.loc[common] - old.loc[common]).abs()
    mad = float(np.median(diff))
    corr = None
    if n >= 8 and new.loc[common].std() > 0 and old.loc[common].std() > 0:
        corr = float(np.corrcoef(new.loc[common], old.loc[common])[0, 1])
    if mad > mad_max:
        return Validation(False, n, mad, corr, f"no coincide con el histórico (MAD {mad:.4g} > {mad_max})")
    if corr is not None and corr < corr_min:
        return Validation(False, n, mad, corr, f"correlación baja con el histórico ({corr:.2f} < {corr_min})")
    return Validation(True, n, mad, corr, "")
