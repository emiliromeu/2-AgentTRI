#!/bin/bash
cd "$(dirname "$0")"

if [ ! -f ".venv/bin/activate" ]; then
    echo "ERROR: no s'ha trobat l'entorn virtual (.venv)."
    echo "Cal crear-lo abans de continuar -- consulta README_APP.md."
    read -p "Prem Enter per tancar..."
    exit 1
fi

source .venv/bin/activate
streamlit run app.py
