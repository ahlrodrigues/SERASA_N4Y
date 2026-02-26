#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$ROOT_DIR/.venv"

log() {
  printf '[%s] %s\n' "$(date '+%H:%M:%S')" "$1"
}

cd "$ROOT_DIR"

if [[ ! -f ".env" ]]; then
  log "ERRO: arquivo .env não encontrado na raiz do projeto."
  exit 1
fi

if [[ ! -d "$VENV_DIR" ]]; then
  log "Criando ambiente virtual em $VENV_DIR"
  python3 -m venv "$VENV_DIR"
fi

log "Ativando ambiente virtual"
# shellcheck disable=SC1091
source "$VENV_DIR/bin/activate"

log "Instalando/atualizando dependencias"
python -m pip install --upgrade pip
pip install -r requirements.txt

log "1/4 - Baixando relatorios SOA"
python Buscar_Relatorio_SOA.py

log "2/4 - Baixando relatorio CNM"
python Buscar_Relatorio_CNM.py

log "3/4 - Baixando relatorio SGP"
# Evita bloqueio no input final do script SGP.
printf '\n' | python Buscar_Relatorio_SGP.py

log "4/4 - Gerando consolidado final"
python -m consolidacao.consolidate

log "Concluido. Arquivo final: consolidacao/output/dashboard_unificado.xlsx"
