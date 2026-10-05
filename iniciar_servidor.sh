#!/usr/bin/env bash
# Script para iniciar el servidor web de viandas
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" >/dev/null 2>&1 && pwd)"
echo "=========================================================="
echo "🍲 INICIANDO SISTEMA DE VIANDAS"
echo "=========================================================="
echo "👉 Página para Clientes:  http://localhost:5000/"
echo "👉 Panel de Cocina:       http://localhost:5000/cocina"
echo "=========================================================="
echo "Presiona Ctrl+C para detener el servidor cuando quieras."
echo ""

"$DIR/venv/bin/python" "$DIR/app.py"
