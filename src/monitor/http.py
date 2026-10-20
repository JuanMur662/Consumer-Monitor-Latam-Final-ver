"""
Cliente HTTP con:
- reintentos y timeouts
- soporte de proxy corporativo y certificados del sistema (truststore)
- copia de cada respuesta en data/raw/<fecha>/ (trazabilidad: qué se descargó, cuándo y de dónde)
- modo --offline: reutiliza la última copia guardada en data/raw/latest/
"""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import logging
import os
import re
import time
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse

import requests

log = logging.getLogger(__name__)

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/128.0 Safari/537.36 MonitorRetailLATAM/1.0")


def _use_system_certificates():
    """En redes corporativas con inspección SSL, usa el almacén de certificados de Windows/macOS."""
    try:
        import truststore  # type: ignore
        truststore.inject_into_ssl()
        return True
    except Exception:
        return False


class OfflineMiss(Exception):
    pass


class Http:
    def __init__(self, raw_dir: Path, offline: bool = False, timeout: int = 60, retries: int = 3,
                 proxy: Optional[str] = None, verify: Optional[str] = None, use_system_certs: bool = True):
        self.raw_dir = Path(raw_dir)
        self.offline = offline
        self.timeout = timeout
        self.retries = retries
        self.today = dt.date.today().isoformat()
        self.day_dir = self.raw_dir / self.today
        self.latest_dir = self.raw_dir / "latest"
        self.down_hosts: dict = {}
        if use_system_certs:
            _use_system_certificates()
        self.s = requests.Session()
        self.s.headers.update({"User-Agent": UA, "Accept-Language": "es-ES,es;q=0.9,en;q=0.8,pt;q=0.7"})
        if proxy:
            self.s.proxies.update({"http": proxy, "https": proxy})
        if verify:
            self.s.verify = verify

    @staticmethod
    def _safe(key: str) -> str:
        return re.sub(r"[^A-Za-z0-9_.-]+", "_", key)[:150]

    @staticmethod
    def redact_text(txt: str) -> str:
        for k, v in os.environ.items():
            if v and len(v) >= 6 and re.search(r"(TOKEN|PASS|KEY|SECRET)", k):
                txt = txt.replace(v, "***")
        return txt

    @staticmethod
    def _redact(obj):
        """Quita tokens/contraseñas antes de escribir el índice de descargas."""
        secrets = [v for k, v in os.environ.items()
                   if v and len(v) >= 6 and re.search(r"(TOKEN|PASS|KEY|SECRET)", k)]
        txt = json.dumps(obj, ensure_ascii=False, default=str)
        for sec in secrets:
            txt = txt.replace(sec, "***")
        txt = re.sub(r'("(?:token|pass|password|api_key)"\s*:\s*)"[^"]*"', r'\1"***"', txt, flags=re.I)
        return json.loads(txt)

    def _save(self, key: str, ext: str, content: bytes, meta: dict):
        meta = self._redact(meta)
        self.day_dir.mkdir(parents=True, exist_ok=True)
        self.latest_dir.mkdir(parents=True, exist_ok=True)
        name = f"{self._safe(key)}.{ext}"
        (self.day_dir / name).write_bytes(content)
        (self.latest_dir / name).write_bytes(content)
        meta = dict(meta, file=name, sha256=hashlib.sha256(content).hexdigest(), bytes=len(content),
                    saved_at=dt.datetime.now().isoformat(timespec="seconds"))
        with open(self.day_dir / "_index.jsonl", "a", encoding="utf-8") as f:
            f.write(json.dumps(meta, ensure_ascii=False) + "\n")

    def fetch(self, key: str, url: str, ext: str = "bin", method: str = "GET", params=None, json_body=None,
              headers=None, save: bool = True, allow_offline: bool = True, reuse_cached: bool = False) -> bytes:
        """Descarga (o lee del caché si offline) y devuelve bytes.

        reuse_cached=True: documentos que no cambian una vez publicados (boletines PDF) se leen de
        data/raw/latest si ya se descargaron en una corrida anterior.
        """
        name = f"{self._safe(key)}.{ext}"
        if reuse_cached and (self.latest_dir / name).exists():
            return (self.latest_dir / name).read_bytes()
        if self.offline:
            p = self.latest_dir / name
            if allow_offline and p.exists():
                return p.read_bytes()
            raise OfflineMiss(f"Sin copia local para {key} ({name})")
        host = urlparse(url).netloc
        if host in self.down_hosts:
            raise RuntimeError(f"{host} no responde en esta corrida (se omite): {self.down_hosts[host]}")
        last_err = None
        for attempt in range(1, self.retries + 1):
            try:
                r = self.s.request(method, url, params=params, json=json_body, headers=headers, timeout=self.timeout)
                if r.status_code >= 500 or r.status_code == 429:
                    raise requests.HTTPError(f"HTTP {r.status_code}")
                if 400 <= r.status_code < 500:
                    # error del cliente (404, 403...): reintentar no ayuda
                    raise PermissionError(f"HTTP {r.status_code} en {url}")
                r.raise_for_status()
                content = r.content
                if save:
                    self._save(key, ext, content, {"key": key, "url": r.url, "method": method,
                                                   "status": r.status_code, "params": params})
                return content
            except PermissionError as e:
                raise RuntimeError(self.redact_text(str(e)))
            except (requests.ConnectionError, requests.Timeout) as e:
                # red bloqueada / host caído: no insistir con el resto de series de ese host
                last_err = e
                if attempt >= 2:
                    self.down_hosts[host] = type(e).__name__
                    break
                time.sleep(1.5)
            except Exception as e:  # noqa: BLE001
                last_err = e
                if attempt < self.retries:
                    time.sleep(1.5 * attempt)
        raise RuntimeError(self.redact_text(f"Fallo descargando {url}: {last_err}"))

    def text(self, key: str, url: str, ext: str = "html", encoding: Optional[str] = None, **kw) -> str:
        b = self.fetch(key, url, ext=ext, **kw)
        if encoding:
            return b.decode(encoding, errors="replace")
        for enc in ("utf-8", "latin-1"):
            try:
                return b.decode(enc)
            except UnicodeDecodeError:
                continue
        return b.decode("utf-8", errors="replace")

    def json(self, key: str, url: str, **kw):
        return json.loads(self.text(key, url, ext="json", **kw))
