"""Orquestación: descargar -> transformar -> validar -> escribir Excel -> noticias -> log."""
from __future__ import annotations

import csv
import datetime as dt
import logging
import os
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import yaml

from . import narrative
from .excel.monitor_sheets import Change, DataTable, Heatmap, norm_label, update_charts
from .excel.sheets_extra import build_log_sheet, build_news_sheet, update_resumen
from .excel.xlsx_book import XlsxBook, quote_sheet
from .http import Http
from .news import google_news, rank_items, rss_feed
from .sources import REGISTRY, apply_transform
from .sources.manual import ensure_template, load_manual
from .validate import validate

log = logging.getLogger("monitor")

# hoja -> (nombre corto, unidad) para la hoja de Noticias
INDICADORES = {
    "Actividad": ("1_Actividad_Economica", "pct", None),
    "Ventas Retail": ("2_Ventas_Retail", "pct", None),
    "ANTAD": ("6_ANTAD", "pct", None),
    "Confianza Empresarial": ("3_Confianza_Empresarial", "nivel", 50),
    "Confianza Consumidor": ("4_Confianza_Consumidor", "nivel", 50),
    "Inflación": ("5_Inflacion", "pct", None),
    "TPM": ("7_TPM's", "pct", None),
}


@dataclass
class Context:
    http: Http
    log: logging.Logger
    diagnostico: bool = False


def _setup_logging(root: Path) -> Path:
    logs = root / "logs"
    logs.mkdir(exist_ok=True)
    path = logs / f"corrida_{dt.datetime.now():%Y%m%d_%H%M%S}.log"
    fmt = logging.Formatter("%(asctime)s %(levelname)-7s %(message)s", "%H:%M:%S")
    root_log = logging.getLogger()
    root_log.handlers.clear()
    root_log.setLevel(logging.INFO)
    fh = logging.FileHandler(path, encoding="utf-8")
    fh.setFormatter(fmt)
    sh = logging.StreamHandler()
    sh.setFormatter(fmt)
    root_log.addHandler(fh)
    root_log.addHandler(sh)
    return path


def _select_and_merge(results):
    """results: lista de (SourceResult, Validation, prioridad). Devuelve segmentos [(res, serie)]."""
    if not results:
        return []

    def quality(t):
        res, v, i = t
        last = res.series.dropna().index.max()
        validated = v.mad is not None
        return (0 if validated else 1, -last.value, v.mad if v.mad is not None else 9e9, i)

    ordered = sorted(results, key=quality)
    base = ordered[0]
    segs = [(base[0], base[0].series.dropna())]
    last = segs[0][1].index.max()
    for res, v, i in sorted(ordered[1:], key=lambda t: -t[0].series.dropna().index.max().value):
        ext = res.series.dropna()
        ext = ext[ext.index > last]
        if not ext.empty:
            segs.append((res, ext))
            last = ext.index.max()
    return segs


def _months_between(a: pd.Timestamp, b: dt.date) -> int:
    return (b.year - a.year) * 12 + (b.month - a.month)


def run(root: Path, offline: bool = False, dry_run: bool = False, solo: Optional[List[str]] = None,
        sin_noticias: bool = False, diagnostico: bool = False) -> int:
    root = Path(root)
    log_path = _setup_logging(root)
    try:
        from dotenv import load_dotenv
        load_dotenv(root / ".env")
    except Exception:  # noqa: BLE001
        pass
    settings = yaml.safe_load(open(root / "config/settings.yaml", encoding="utf-8"))
    cfg = yaml.safe_load(open(root / "config/series.yaml", encoding="utf-8"))
    news_cfg = yaml.safe_load(open(root / "config/noticias.yaml", encoding="utf-8"))
    started = dt.datetime.now()
    run_info = {"inicio": started.strftime("%d-%b-%Y %H:%M"), "modo": "offline" if offline else "en línea",
                "dias": news_cfg.get("dias", 45)}
    log.info("Monitor Retail LATAM: inicio (%s). Log: %s", run_info["modo"], log_path)

    excel = root / settings["excel"]
    if not excel.exists():
        orig = excel.with_name(excel.stem + "_original.xlsx")
        if orig.exists():
            shutil.copy2(orig, excel)
            log.info("Primera corrida: se creó %s a partir de la plantilla original", excel.name)
        else:
            raise FileNotFoundError(f"No existe el Excel {excel}")

    http = Http(root / settings["raw"], offline=offline, timeout=settings.get("timeout_segundos", 60),
                retries=settings.get("reintentos", 3), proxy=settings.get("proxy") or None,
                use_system_certs=settings.get("certificados_sistema", True))
    ctx = Context(http, log, diagnostico)
    manual_path = root / settings["manual"]
    ensure_template(manual_path)
    manual = load_manual(manual_path)

    book = XlsxBook(str(excel))
    status_rows: List[dict] = []
    status_by_sheet: Dict[str, Dict[str, str]] = {}
    all_changes: List[Change] = []
    tables: Dict[str, DataTable] = {}

    for hoja, hcfg in cfg["hojas"].items():
        if solo and hoja not in solo:
            continue
        try:
            table = DataTable.locate(book, hoja)
        except Exception as e:  # noqa: BLE001
            log.error("Hoja %s: %s", hoja, e)
            continue
        table.normalize_formula_dates()
        tables[hoja] = table
        vcfg = hcfg.get("validacion", {})
        for col_label, scfg in hcfg["series"].items():
            if norm_label(col_label) not in table.columns:
                log.warning("%s: columna %s no existe en la hoja", hoja, col_label)
                continue
            existing = table.read(col_label)
            attempts, results = [], []
            for prio, src in enumerate(scfg.get("fuentes") or []):
                tipo = src.get("tipo")
                req = src.get("requiere")
                if req and not os.environ.get(req, "").strip():
                    attempts.append(f"{tipo}: omitida (falta {req} en .env)")
                    continue
                try:
                    res = REGISTRY[tipo](src, ctx)
                    res.raw = res.series
                    res.series = apply_transform(res.series, src.get("transform"))
                    vargs = (float(src.get("mad_max", vcfg.get("mad_max", 0.01))), float(vcfg.get("corr_min", 0.6)),
                             tuple(vcfg.get("rango", (None, None))))
                    v = validate(res.series, existing, *vargs)
                    if res.alternatives and not (v.ok and v.mad is not None):
                        # archivo con varias filas candidatas: quedarse con la que coincide con el histórico
                        best = None
                        for lab, alt in res.alternatives:
                            try:
                                alt_t = apply_transform(alt, src.get("transform"))
                            except Exception:  # noqa: BLE001
                                continue
                            va = validate(alt_t, existing, *vargs)
                            if va.ok and va.mad is not None and (best is None or va.mad < best[2].mad):
                                best = (lab, alt_t, va)
                        if best:
                            res.title, res.series, v = best[0], best[1], best[2]
                            attempts.append(f"{res.source}: se eligió la fila «{best[0]}» por coincidir con el histórico")
                    last = res.series.dropna().index.max()
                    etiqueta = f" [{res.title[:70]}]" if res.title and tipo in ("archivo_web", "inegi") else ""
                    attempts.append(f"{res.source}{etiqueta}: {'OK' if v.ok else 'RECHAZADA'} hasta {last:%Y-%m} ({v.short()})")
                    if v.ok:
                        results.append((res, v, prio))
                except Exception as e:  # noqa: BLE001
                    attempts.append(f"{tipo}: ERROR {Http.redact_text(str(e))[:220]}")
            segs = _select_and_merge(results)
            n_written, changes, used, estado, val_txt = 0, [], [], "ERROR", ""
            if segs:
                base_v = next(v for r, v, i in results if r is segs[0][0])
                val_txt = base_v.short()
                estado = "OK" if all(r.verified for r, _ in segs) else "NOTICIAS"
                used = [r.source for r, _ in segs]
            if not settings.get("escribir_historia_completa", True):
                # solo meses que aún no tienen dato en el Excel (no reescribe revisiones)
                have = existing.dropna().index
                segs = [(r, s_[~s_.index.isin(have)]) for r, s_ in segs]
                segs = [(r, s_) for r, s_ in segs if not s_.empty]
            for res, s in segs:
                ch = table.write_series(col_label, s, res.source, verified=res.verified)
                changes += ch
                n_written += len(s)
            # manual: prioridad sobre lo automático para el mes indicado
            man = manual.get((hoja, norm_label(col_label)))
            if man:
                for ts, m in sorted(man.items()):
                    ch = table.write_series(col_label, pd.Series([m["valor"]], index=[ts]),
                                            f"MANUAL: {m['fuente']}", verified=m["verificado"])
                    changes += ch
                used.append("manual_inputs.csv")
                if estado == "ERROR":
                    estado = "MANUAL"
            if not (scfg.get("fuentes") or []) and not man:
                estado = "SIN FUENTE"
                attempts.append("sin fuente automática configurada: usar data/manual_inputs.csv")
            all_changes += changes
            last_series = table.read(col_label)
            ultimo = last_series.index.max().strftime("%Y-%m") if not last_series.empty else ""
            nuevos = sum(1 for c in changes if c.kind == "nuevo")
            if estado == "OK" and not changes:
                detalle_extra = "sin datos nuevos ni revisiones"
            else:
                detalle_extra = f"{nuevos} nuevos, {len(changes) - nuevos} revisados"
            status_rows.append({"hoja": hoja, "serie": col_label, "estado": estado, "fuente": " + ".join(used),
                                "ultimo": ultimo, "escritos": n_written, "validacion": val_txt,
                                "detalle": detalle_extra + " | " + " || ".join(attempts)})
            status_by_sheet.setdefault(hoja, {})[norm_label(col_label)] = estado
            lvl = logging.INFO if estado in ("OK", "MANUAL", "NOTICIAS") else logging.WARNING
            log.log(lvl, "%-24s %-15s %-10s último=%s  %s", hoja, col_label, estado, ultimo, detalle_extra)
            for a in attempts:
                log.info("      · %s", a)

        # heatmap, gráficos y nombres
        hm = Heatmap.locate(table.sheet)
        if hm:
            end = hm.refresh(table, settings.get("heatmap_promedio_ultimos", 12))
            book.set_defined_name(f"{hcfg['alias']}_Heatmap", hm.defined_name_ref())
            log.info("%s: heatmap hasta %s", hoja, end)
        n = update_charts(book, table)
        book.set_defined_name(f"{hcfg['alias']}_Datos", table.defined_name_ref(),
                              "Tabla de input; la actualiza el código Python")
        last_row = table.last_data_row()
        book.set_defined_name(f"{hcfg['alias']}_Fechas",
                              f"{quote_sheet(hoja)}!${table.date_col}${table.first_row}:${table.date_col}${last_row}")
        for key, col in table.columns.items():
            nm = f"{hcfg['alias']}_{key.replace(' ', '_')}"
            book.set_defined_name(nm, f"{quote_sheet(hoja)}!${col}${table.first_row}:${col}${last_row}")
        log.info("%s: %d rangos de gráficos ajustados hasta la fila %d", hoja, n, last_row)

    # ------------------------------------------------------------------ noticias
    if settings.get("noticias", True) and not sin_noticias:
        try:
            blocks = build_news_blocks(ctx, book, tables, news_cfg, settings)
            build_news_sheet(book, blocks, run_info)
            log.info("Hoja Noticias: %d países", len(blocks))
        except Exception as e:  # noqa: BLE001
            log.error("No se pudo construir la hoja Noticias: %s", e)

    ok = sum(1 for r in status_rows if r["estado"] == "OK")
    run_info.update({"ok": ok, "total": len(status_rows),
                     "nuevos": sum(1 for c in all_changes if c.kind == "nuevo"),
                     "revisiones": sum(1 for c in all_changes if c.kind == "revision")})
    build_log_sheet(book, status_rows, all_changes, run_info)
    update_resumen(book, run_info, status_by_sheet)
    book.set_active_sheet("Resumen")
    book.set_full_calc_on_load()
    book.drop_calc_chain()

    # ------------------------------------------------------------------ guardar
    stamp = started.strftime("%Y%m%d_%H%M")
    rep_dir = root / settings["reportes"]
    rep_dir.mkdir(parents=True, exist_ok=True)
    if dry_run:
        out = rep_dir / f"PREVIEW_{excel.stem}_{stamp}.xlsx"
        book.save(str(out))
        log.info("Modo prueba: NO se tocó el Excel maestro. Vista previa: %s", out)
    else:
        bk_dir = root / settings["respaldos"]
        bk_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(excel, bk_dir / f"{excel.stem}_antes_{stamp}.xlsx")
        tmp = excel.with_suffix(".tmp.xlsx")
        book.save(str(tmp))
        try:
            os.replace(tmp, excel)
            out = excel
        except PermissionError:
            out = excel.with_name(f"{excel.stem}_actualizado_{stamp}.xlsx")
            os.replace(tmp, out)
            log.warning("El Excel está abierto (bloqueado). Guardé la versión nueva como %s", out.name)
        log.info("Excel actualizado: %s", out)

    # reportes CSV (trazabilidad)
    with open(rep_dir / f"estado_{stamp}.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(status_rows[0].keys()) if status_rows else ["hoja"])
        w.writeheader()
        w.writerows(status_rows)
    with open(rep_dir / f"cambios_{stamp}.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["hoja", "serie", "mes", "celda", "anterior", "nuevo", "tipo", "fuente"])
        for c in all_changes:
            w.writerow([c.sheet, c.column, c.month, c.cell, c.old, c.new, c.kind, c.source])
    store = root / settings["store"]
    store.mkdir(parents=True, exist_ok=True)
    for hoja, t in tables.items():
        df = pd.DataFrame({t.labels[k]: t.read(t.labels[k]) for k in t.columns})
        df.index.name = "mes"
        df.to_csv(store / f"{hoja.replace(chr(39), '')}.csv", float_format="%.6f")
    log.info("Resumen: %d/%d series OK | %d datos nuevos | %d revisiones | reporte en %s",
             ok, len(status_rows), run_info["nuevos"], run_info["revisiones"], rep_dir)
    return 0 if ok > 0 else 1


def build_news_blocks(ctx, book: XlsxBook, tables: Dict[str, DataTable], news_cfg: dict, settings: dict) -> list:
    today = dt.date.today()
    blocks = []
    feeds = news_cfg.get("feeds") or []
    use_llm = str(settings.get("resumen_llm", "auto")).lower() not in ("false", "no", "0")
    for pais, pcfg in (news_cfg.get("paises") or {}).items():
        inds = []
        for ind_name in pcfg.get("indicadores", {}):
            hoja, unidad, umbral = INDICADORES.get(ind_name, (None, None, None))
            col = "OVERALL" if ind_name == "ANTAD" else pais
            t = tables.get(hoja)
            if t is None or norm_label(col) not in t.columns:
                continue
            s = t.read(col).dropna()
            if s.empty:
                continue
            last = s.index.max()
            prev = s.get(last - pd.DateOffset(months=1))
            y12 = s.get(last - pd.DateOffset(months=12))
            ind = {"nombre": ind_name, "mes": narrative.mes_txt(last), "valor": float(s[last]),
                   "previo": None if prev is None else float(prev), "hace12": None if y12 is None else float(y12),
                   "unidad": unidad, "umbral": umbral, "rezago_meses": _months_between(last, today) - 1}
            ind["delta"] = None if ind["previo"] is None else ind["valor"] - ind["previo"]
            ind["lectura"] = narrative.lectura(ind)
            inds.append(ind)
        items = []
        for ind_name, q in (pcfg.get("indicadores") or {}).items():
            try:
                got = google_news(ctx, q["consulta"], pcfg.get("idioma", "es-419"), pcfg.get("pais", "CO"),
                                  news_cfg.get("dias", 45))
                top = rank_items(got, q.get("filtro"), q.get("excluir"), news_cfg.get("max_por_indicador", 3),
                                 news_cfg.get("dias", 45) + 5)
                for it in top:
                    it.tags = [ind_name]
                items += top
            except Exception as e:  # noqa: BLE001
                ctx.log.warning("Noticias %s/%s: %s", pais, ind_name, e)
        for fd in feeds:
            if fd.get("pais") != pais:
                continue
            try:
                got = rss_feed(ctx, fd["url"], fd["nombre"])
                top = rank_items(got, fd.get("filtro"), None, 2, news_cfg.get("dias", 45) + 5)
                for it in top:
                    it.tags = [fd.get("indicador", "Oficial")]
                items += top
            except Exception as e:  # noqa: BLE001
                ctx.log.warning("Feed %s: %s", fd.get("nombre"), e)
        texto = None
        if use_llm:
            texto = narrative.llm_summary(pais, inds, items, settings.get("modelo_llm", "claude-sonnet-5-5"))
        if not texto:
            texto = narrative.rule_based(pais, inds, items)
        blocks.append({"pais": pais, "narrativa": texto, "indicadores": inds, "noticias": items})
    return blocks
