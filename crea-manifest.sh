#!/bin/bash
# Crea manifest.xml con l'indirizzo dove hai pubblicato i file.
# Uso:  ./crea-manifest.sh https://TUO-UTENTE.github.io/outlook-traduttore
set -e
BASE="${1%/}"
if [[ ! "$BASE" =~ ^https:// ]]; then
  echo "Indica l'indirizzo HTTPS, es.: ./crea-manifest.sh https://riccardo.github.io/outlook-traduttore"
  exit 1
fi
cd "$(dirname "$0")"
sed "s#__BASE_URL__#${BASE}#g" manifest.template.xml > manifest.xml
echo "Creato manifest.xml per ${BASE}"
