"""
Datos manuales: data/manual_inputs.csv

Columnas: hoja, columna, mes (AAAA-MM), valor, unidad (pct|fraccion|nivel), fuente, verificado (si|no), nota

- `pct`: el valor viene en porcentaje (2.3 = 2.3%) y se guarda como 0.023
- `fraccion`: ya viene como 0.023
- `nivel`: índices (confianza 0-100)
Los datos manuales tienen prioridad sobre las fuentes automáticas para ese mes.
"""
from __future__ import annotations

import csv
from pathlib import Path
from typing import Dict, Tuple

import pandas as pd

from .base import parse_number, parse_period

HEADER = ["hoja", "columna", "mes", "valor", "unidad", "fuente", "verificado", "nota"]


def load_manual(path: Path) -> Dict[Tuple[str, str], Dict[pd.Timestamp, dict]]:
    out: Dict[Tuple[str, str], Dict[pd.Timestamp, dict]] = {}
    if not path.exists():
        return out
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    for row in rows:
        if not row.get("hoja") or str(row.get("hoja")).startswith("#"):
            continue
        ts = parse_period(row.get("mes"))
        v = parse_number(row.get("valor"))
        if ts is None or v is None:
            continue
        unidad = (row.get("unidad") or "fraccion").strip().lower()
        if unidad == "pct":
            v = v / 100.0
        key = (row["hoja"].strip(), row["columna"].strip().upper())
        out.setdefault(key, {})[ts] = {
            "valor": v,
            "fuente": (row.get("fuente") or "manual").strip(),
            "verificado": (row.get("verificado") or "").strip().lower() in ("si", "sí", "yes", "true", "1", "x"),
            "nota": row.get("nota") or "",
        }
    return out


def ensure_template(path: Path):
    if path.exists():
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(HEADER)
        w.writerow(["# ejemplo (las filas que empiezan con # se ignoran)", "OVERALL", "2026-09", "2.3", "pct",
                    "Comunicado ANTAD", "si", ""])
