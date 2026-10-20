#!/usr/bin/env bash
# Monitor Retail LATAM - instalación en macOS / Linux
set -e
cd "$(dirname "$0")"
PY=$(command -v python3 || command -v python)
[ -x .venv/bin/python ] || "$PY" -m venv .venv
.venv/bin/python -m pip install --upgrade pip -q
.venv/bin/python -m pip install -r requirements.txt
[ -f .env ] || cp .env.example .env
echo "Listo. Ejecuta ./run.sh"
