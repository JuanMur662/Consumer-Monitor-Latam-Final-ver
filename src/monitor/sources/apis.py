"""
Fuentes con API pública estable (JSON/CSV).

BCB SGS (Brasil), IBGE SIDRA (Brasil), BCRP (Perú), BIS (todos), FRED (EE.UU.),
Banxico SIE (México, token), INEGI BIE (México, token), BCCh (Chile, API con token
o tablas públicas de la BDE sin token).
"""
from __future__ import annotations

import datetime as dt
import io
import json
import os
import re
from typing import Any, Dict, List

import pandas as pd
from bs4 import BeautifulSoup

from .base import SourceResult, parse_number, parse_period, register, to_monthly, strip_accents

START_YEAR = 2018


def _need_env(var: str) -> str:
    v = os.environ.get(var, "").strip()
    if not v:
        raise RuntimeError(f"Falta {var} en el archivo .env (ver README: tokens gratuitos)")
    return v


# ---------------------------------------------------------------------------- BCB SGS
@register("bcb_sgs")
def bcb_sgs(spec: Dict[str, Any], ctx) -> SourceResult:
    code = spec["codigo"]
    today = dt.date.today()
    # las series diarias tienen límite de 10 años por consulta
    start = dt.date(max(START_YEAR, today.year - 9), 1, 1)
    url = (f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{code}/dados?formato=json"
           f"&dataInicial={start:%d/%m/%Y}&dataFinal={today:%d/%m/%Y}")
    data = ctx.http.json(f"bcb_sgs_{code}", url)
    pts = []
    for row in data:
        d = dt.datetime.strptime(row["data"], "%d/%m/%Y")
        pts.append((d, parse_number(row["valor"], ".")))
    s = to_monthly(pts, spec.get("agregacion", "last"))
    return SourceResult(s, f"BCB SGS {code}", f"https://www3.bcb.gov.br/sgspub/ (serie {code})")


# ---------------------------------------------------------------------------- IBGE SIDRA
@register("ibge_sidra")
def ibge_sidra(spec: Dict[str, Any], ctx) -> SourceResult:
    t, v = spec["tabela"], spec["variavel"]
    cls = spec.get("classificacao", "")  # p. ej. "c11046/56734"
    url = f"https://apisidra.ibge.gov.br/values/t/{t}/n1/all/v/{v}/p/all" + (f"/{cls}" if cls else "")
    data = ctx.http.json(f"ibge_sidra_{t}_{v}_{cls}", url)
    header, rows = data[0], data[1:]
    per_key = None
    for k, lab in header.items():
        if k.endswith("C") and ("Mês" in lab or "Mes" in lab) and "Código" in lab:
            per_key = k
    if per_key is None:
        per_key = "D3C"
    pts = [(parse_period(r[per_key]), parse_number(r["V"], ".")) for r in rows]
    s = to_monthly(pts)
    return SourceResult(s, f"IBGE SIDRA t{t} v{v}", f"https://sidra.ibge.gov.br/tabela/{t}")


# ---------------------------------------------------------------------------- BCRP
_BCRP_M = {"Ene": 1, "Feb": 2, "Mar": 3, "Abr": 4, "May": 5, "Jun": 6, "Jul": 7, "Ago": 8,
           "Sep": 9, "Set": 9, "Oct": 10, "Nov": 11, "Dic": 12}


@register("bcrp")
def bcrp(spec: Dict[str, Any], ctx) -> SourceResult:
    code = spec["codigo"]
    today = dt.date.today()
    url = (f"https://estadisticas.bcrp.gob.pe/estadisticas/series/api/{code}/json/"
           f"{START_YEAR}-1/{today.year}-12")
    data = json.loads(ctx.http.text(f"bcrp_{code}", url, ext="json"))
    pts = []
    for p in data.get("periods", []):
        name = p["name"]
        m = re.match(r"([A-Za-z]{3})\.(\d{4})", name)
        if not m:
            continue
        mo = _BCRP_M.get(m.group(1).capitalize())
        if not mo:
            continue
        pts.append((pd.Timestamp(int(m.group(2)), mo, 1), parse_number(p["values"][0], ".")))
    s = to_monthly(pts)
    title = data.get("config", {}).get("series", [{}])[0].get("name", "")
    return SourceResult(s, f"BCRP {code}", f"https://estadisticas.bcrp.gob.pe/estadisticas/series/mensuales/resultados/{code}/html", title)


# ---------------------------------------------------------------------------- BIS
@register("bis")
def bis(spec: Dict[str, Any], ctx) -> SourceResult:
    flow, key = spec["dataset"], spec["clave"]
    url = f"https://stats.bis.org/api/v1/data/{flow}/{key}/all?startPeriod={START_YEAR}-01&format=csv"
    txt = ctx.http.text(f"bis_{flow}_{key}", url, ext="csv")
    df = pd.read_csv(io.StringIO(txt))
    pts = [(parse_period(p), parse_number(v, ".")) for p, v in zip(df["TIME_PERIOD"], df["OBS_VALUE"])]
    s = to_monthly(pts)
    return SourceResult(s, f"BIS {flow} {key}", f"https://data.bis.org/topics/{'CBPOL' if 'CBPOL' in flow else 'CPI'}")


# ---------------------------------------------------------------------------- FRED
@register("fred")
def fred(spec: Dict[str, Any], ctx) -> SourceResult:
    code = spec["codigo"]
    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={code}"
    txt = ctx.http.text(f"fred_{code}", url, ext="csv")
    df = pd.read_csv(io.StringIO(txt))
    dcol, vcol = df.columns[0], df.columns[1]   # 'observation_date' (antes 'DATE') y el código
    pts = [(parse_period(d), parse_number(v, ".")) for d, v in zip(df[dcol], df[vcol])]
    s = to_monthly(pts, spec.get("agregacion", "last"))
    return SourceResult(s, f"FRED {code}", f"https://fred.stlouisfed.org/series/{code}")


# ---------------------------------------------------------------------------- Banxico SIE
@register("banxico")
def banxico(spec: Dict[str, Any], ctx) -> SourceResult:
    sid = spec["serie"]
    token = _need_env("BANXICO_TOKEN")
    today = dt.date.today()
    url = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{sid}/datos/{START_YEAR}-01-01/{today:%Y-%m-%d}"
    data = ctx.http.json(f"banxico_{sid}", url, headers={"Bmx-Token": token, "Accept": "application/json"})
    serie = data["bmx"]["series"][0]
    pts = [(dt.datetime.strptime(d["fecha"], "%d/%m/%Y"), parse_number(d["dato"], ".")) for d in serie.get("datos", [])]
    s = to_monthly(pts, spec.get("agregacion", "last"))
    return SourceResult(s, f"Banxico SIE {sid}", "https://www.banxico.org.mx/SieInternet/", serie.get("titulo", ""))


# ---------------------------------------------------------------------------- INEGI BIE
INEGI_EXPORT = "https://www.inegi.org.mx/app/api/indicadores/interna_v1_3/API.svc/ExportacionBancoInformacion"
INEGI_SEARCH = "https://www.inegi.org.mx/app/api/buscadorcore/v1/busquedaBIE/"
INEGI_DEV = "https://www.inegi.org.mx/app/api/indicadores/desarrolladores/jsonxml/INDICATOR/{id}/es/{area}/false/{bank}/2.0/{token}?type=json"


def inegi_search(ctx, query: str) -> List[Dict[str, str]]:
    body = {"busqueda": query, "busquedaCiencia": "", "paginaInicio": 0, "paginaFin": 40,
            "filtrobusqueda": "CBUSQUEDA", "filtrotema": "null", "orderby": "RANKING", "orderbyAscDesc": "Desc",
            "metodoBusqueda": 1, "herramienta": 32}
    key = "inegi_busqueda_" + re.sub(r"\W+", "_", strip_accents(query))[:60]
    raw = ctx.http.fetch(key, INEGI_SEARCH, ext="json", method="POST", json_body=body,
                         headers={"Content-Type": "application/json; charset=utf-8", "Accept": "application/json"})
    data = json.loads(raw.decode("utf-8"))
    out = []
    for r in data if isinstance(data, list) else []:
        out.append({"id": str(r.get("INDICADOR", "")),
                    "titulo": re.sub(r"\s+", " ", str(r.get("TITULO", "")).replace("#null", "")).strip()})
    return out


def _inegi_points_export(ctx, ind: str, token: str):
    body = {"areasGeograficas": "null", "casoExportacion": "indicadorVertical", "fechaInicio": str(START_YEAR),
            "fechaFin": str(dt.date.today().year + 1), "formato": "json", "idioma": "es", "indicadores": ind,
            "mostrarDecimales": "true", "mostrarEstadistico": "false", "ordenaPeriodo": "ap", "orden": "a",
            "tematica": "3", "token": token}
    raw = ctx.http.fetch(f"inegi_export_{ind}", INEGI_EXPORT, ext="json", method="POST", json_body=body,
                         headers={"Content-Type": "application/json; charset=utf-8", "Accept": "application/json"})
    data = json.loads(raw.decode("utf-8"))
    if not isinstance(data, list):
        raise RuntimeError(f"INEGI export sin datos: {str(data)[:200]}")
    rows = [r.get("listaCeldas", []) for r in data]
    header = rows[0]
    col = next(i for i, c in enumerate(header) if i > 0 and str(c.get("valor", "")).strip() == ind)
    pts = []
    for cells in rows[1:]:
        per = str(cells[0].get("valor", "")).replace("\\", "")
        pts.append((parse_period(per), parse_number(str(cells[col].get("valor", "")).replace(",", ""), ".")))
    return pts


def _inegi_points_dev(ctx, ind: str, token: str):
    last = None
    for bank, area in (("BIE-BISE", "0700"), ("BIE", "0700"), ("BIE-BISE", "00")):
        try:
            url = INEGI_DEV.format(id=ind, area=area, bank=bank, token=token)
            data = ctx.http.json(f"inegi_dev_{ind}_{bank}_{area}", url)
            obs = data["Series"][0]["OBSERVATIONS"]
            return [(parse_period(o["TIME_PERIOD"]), parse_number(o["OBS_VALUE"], ".")) for o in obs]
        except Exception as e:  # noqa: BLE001
            last = e
    raise RuntimeError(f"INEGI API desarrolladores: {last}")


@register("inegi")
def inegi(spec: Dict[str, Any], ctx) -> SourceResult:
    """Indicador del BIE. Si se da `buscar` + `titulo_regex`, el id se resuelve (y verifica) por título."""
    token = _need_env("INEGI_TOKEN")
    ind = str(spec.get("id") or "")
    title = ""
    if spec.get("buscar"):
        rx = re.compile(strip_accents(spec.get("titulo_regex", ".*")), re.I)
        try:
            hits = inegi_search(ctx, spec["buscar"])
        except Exception as e:  # noqa: BLE001
            if not ind:
                raise RuntimeError(f"INEGI búsqueda falló y no hay id configurado: {e}")
            hits = []
        xrx = re.compile(strip_accents(spec["excluir_regex"]), re.I) if spec.get("excluir_regex") else None
        match = [h for h in hits if rx.search(strip_accents(h["titulo"]))
                 and not (xrx and xrx.search(strip_accents(h["titulo"])))]
        if ind and any(h["id"] == ind for h in match):
            title = next(h["titulo"] for h in match if h["id"] == ind)
        elif match:
            ind, title = match[0]["id"], match[0]["titulo"]
        elif not ind:
            raise RuntimeError("INEGI: la búsqueda no encontró un indicador que cumpla titulo_regex")
    errors = []
    for fn in (_inegi_points_export, _inegi_points_dev):
        try:
            pts = fn(ctx, ind, token)
            s = to_monthly(pts)
            if not s.empty:
                return SourceResult(s, f"INEGI BIE {ind}", f"https://www.inegi.org.mx/app/indicadores/?ind={ind}", title)
        except Exception as e:  # noqa: BLE001
            errors.append(f"{fn.__name__}: {e}")
    raise RuntimeError("; ".join(errors))


# ---------------------------------------------------------------------------- BCCh
_SIETE = "https://si3.bcentral.cl/Siete/ES/Siete/Cuadro/{cuadro}"


@register("bcch_siete")
def bcch_siete(spec: Dict[str, Any], ctx) -> SourceResult:
    """Tabla pública de la Base de Datos Estadísticos del BCCh (no requiere token)."""
    cuadro = spec["cuadro"]
    params = {"cbFechaInicio": str(START_YEAR), "cbFechaTermino": str(dt.date.today().year),
              "cbFrecuencia": "MONTHLY", "cbCalculo": spec.get("calculo", "NONE"), "cbFechaBase": ""}
    html = ctx.http.text("bcch_siete_" + cuadro.replace("/", "_") + "_" + params["cbCalculo"],
                         _SIETE.format(cuadro=cuadro), params=params)
    soup = BeautifulSoup(html, "lxml")
    rx = re.compile(strip_accents(spec["fila"]), re.I)
    for table in soup.find_all("table"):
        trs = table.find_all("tr")
        header_idx, periods = None, {}
        for i, tr in enumerate(trs[:6]):
            cells = [c.get_text(" ", strip=True) for c in tr.find_all(["th", "td"])]
            per = {j: parse_period(t) for j, t in enumerate(cells)}
            per = {j: p for j, p in per.items() if p is not None}
            if len(per) >= 6:
                header_idx, periods = i, per
                break
        if header_idx is None:
            continue
        for tr in trs[header_idx + 1:]:
            cells = [c.get_text(" ", strip=True) for c in tr.find_all(["th", "td"])]
            label = next((t for t in cells[:3] if t and not re.fullmatch(r"[\d.,\-\s]*", t)), "")
            if not rx.search(strip_accents(label)):
                continue
            # alinear por la derecha: algunas filas tienen una celda de selección menos que el encabezado
            offset = len(trs[header_idx].find_all(["th", "td"])) - len(cells)
            pts = []
            for j, p in periods.items():
                k = j - offset
                if 0 <= k < len(cells):
                    pts.append((p, parse_number(cells[k], ",")))
            s = to_monthly(pts)
            if not s.empty:
                return SourceResult(s, f"BCCh BDE {cuadro.split('/')[-1]}", _SIETE.format(cuadro=cuadro), label)
    raise RuntimeError(f"BCCh: no encontré la fila '{spec['fila']}' en el cuadro {cuadro}")


@register("bcch_api")
def bcch_api(spec: Dict[str, Any], ctx) -> SourceResult:
    code = spec["codigo"]
    params = {"function": "GetSeries", "timeseries": code, "firstdate": f"{START_YEAR}-01-01"}
    tok = os.environ.get("BCCH_TOKEN", "").strip()
    if tok:
        params["token"] = tok
    else:
        params["user"] = _need_env("BCCH_USER")
        params["pass"] = _need_env("BCCH_PASS")
    raw = ctx.http.fetch(f"bcch_api_{code}", "https://si3.bcentral.cl/SieteRestWS/SieteRestWS.ashx",
                         ext="json", params=params, save=True)
    data = json.loads(raw.decode("utf-8"))
    if data.get("Codigo") != 0:
        raise RuntimeError(f"BCCh API: {data.get('Descripcion')}")
    obs = data["Series"]["Obs"]
    pts = [(dt.datetime.strptime(o["indexDateString"], "%d-%m-%Y"), parse_number(o["value"], "."))
           for o in obs if o.get("statusCode", "OK") == "OK"]
    s = to_monthly(pts, spec.get("agregacion", "last"))
    return SourceResult(s, f"BCCh API {code}", "https://si3.bcentral.cl/siete", data["Series"].get("descripEsp", ""))
