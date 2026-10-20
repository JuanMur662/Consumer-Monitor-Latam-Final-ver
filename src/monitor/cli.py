"""
Uso (desde la carpeta del proyecto, con el entorno virtual activo):

    python -m monitor                 # corrida normal: descarga, valida, actualiza el Excel y noticias
    python -m monitor --prueba        # igual, pero guarda una VISTA PREVIA y no toca el Excel maestro
    python -m monitor --offline       # reutiliza lo último descargado en data/raw/latest (sin internet)
    python -m monitor --sin-noticias  # omite la hoja Noticias
    python -m monitor --solo 5_Inflacion "7_TPM's"
    python -m monitor --diagnostico   # imprime la estructura de los archivos que no se pudieron leer
    python -m monitor buscar-inegi "confianza empresarial comercio"   # encuentra ids del INEGI
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path


def project_root() -> Path:
    here = Path(__file__).resolve()
    for p in [here.parent.parent.parent, Path.cwd()]:
        if (p / "config" / "series.yaml").exists():
            return p
    return Path.cwd()


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="monitor", description="Actualiza el Monitor Retail LATAM")
    ap.add_argument("comando", nargs="?", default="actualizar", choices=["actualizar", "buscar-inegi"])
    ap.add_argument("texto", nargs="?", default="")
    ap.add_argument("--prueba", action="store_true", help="no modifica el Excel maestro; genera vista previa")
    ap.add_argument("--offline", action="store_true", help="usa las últimas descargas guardadas")
    ap.add_argument("--sin-noticias", action="store_true")
    ap.add_argument("--diagnostico", action="store_true")
    ap.add_argument("--solo", nargs="*", help="hojas a procesar")
    a = ap.parse_args(argv)
    root = project_root()
    if a.comando == "buscar-inegi":
        return buscar_inegi(root, a.texto)
    from .pipeline import run
    return run(root, offline=a.offline, dry_run=a.prueba, solo=a.solo, sin_noticias=a.sin_noticias,
               diagnostico=a.diagnostico)


def buscar_inegi(root: Path, texto: str) -> int:
    from dotenv import load_dotenv
    from .http import Http
    from .pipeline import Context
    from .sources.apis import inegi_search
    import logging
    load_dotenv(root / ".env")
    ctx = Context(Http(root / "data/raw"), logging.getLogger("monitor"))
    for h in inegi_search(ctx, texto)[:40]:
        print(f"{h['id']:>10}  {h['titulo']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
