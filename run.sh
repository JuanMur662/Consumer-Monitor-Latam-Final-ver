#!/usr/bin/env bash
cd "$(dirname "$0")"
[ -x .venv/bin/python ] || ./setup.sh
PYTHONPATH="$(pwd)/src" .venv/bin/python -m monitor "$@"
