#!/usr/bin/env bash
# Script de acceso rápido para el sistema de viandas
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" >/dev/null 2>&1 && pwd)"
"$DIR/venv/bin/python" "$DIR/viandas_engine.py" "$@"
