#!/bin/bash
# =============================================================================
#  "Il pulsante" — macOS. Doppio clic dal Finder, oppure:  ./run_report.command 31
#
#  Prima volta, in questa cartella:
#      python3 -m pip install -e ".[excel]"
#      chmod +x run_report.command
#  Serve Excel desktop: il ricalcolo delle formule ad array 365 e del VBA non
#  e' replicabile altrove.
# =============================================================================
set -u
cd "$(dirname "$0")"

WEEK="${1:-}"
if [ -z "$WEEK" ]; then
    read -r -p "Numero settimana (es. 31): " WEEK
fi
if [ -z "$WEEK" ]; then
    echo "Nessuna settimana indicata. Esco."
    exit 2
fi

echo
echo "=== Controllo dei CSV in input/ ==="
if ! python3 -m fasterreports.omni.cli preflight --week "$WEEK"; then
    echo
    echo "I CSV non sono a posto: vedi i punti BLOCCATO qui sopra."
    echo "Nessun workbook e' stato prodotto."
    read -r -p "Premi Invio per chiudere."
    exit 1
fi

echo
echo "=== Generazione del report ==="
if ! python3 -m fasterreports.omni.cli build --week "$WEEK"; then
    echo
    echo "Generazione fallita."
    read -r -p "Premi Invio per chiudere."
    exit 1
fi

echo
echo "Fatto. Il file e' in output/"
read -r -p "Premi Invio per chiudere."
