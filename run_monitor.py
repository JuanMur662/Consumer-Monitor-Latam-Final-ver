"""Atajo para VS Code: abre este archivo y presiona "Run Python File" (▶).

Equivale a `python -m monitor` usando el código de la carpeta src/.
Para pasar opciones: python run_monitor.py --prueba
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from monitor.cli import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
