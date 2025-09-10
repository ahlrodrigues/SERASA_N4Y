# consolidacao/config.py
import os
from pathlib import Path

# Estilo visual dos status no Excel:
# - "UPPER": NEGATIVADO / BAIXADO / PENDENTE / ERRO / DETERMINAÇÃO
# - "SOA"  : Negativado / Baixado / Pendente / erro / Determinação
CONSOLIDADO_STYLE = "UPPER"

BASE_DIR = Path(__file__).resolve().parent
DOWNLOAD_DIR = Path(os.getenv("N4Y_DOWNLOAD_DIR", BASE_DIR / "download")).resolve()
OUTPUT_DIR   = Path(os.getenv("N4Y_OUTPUT_DIR",   BASE_DIR / "output")).resolve()

# Arquivos de entrada
CNM_STD_XLSX = DOWNLOAD_DIR / "Relatorio_CNM_standard.xlsx"
CNM_RAW_XLSX = DOWNLOAD_DIR / "Relatorio_CNM.xlsx"
SGP_XLSX     = DOWNLOAD_DIR / "Relatorio_SGP.xlsx"
SOA_CSVS     = {
    "Ativas":       DOWNLOAD_DIR / "Ativas.csv",
    "Baixadas":     DOWNLOAD_DIR / "Baixadas.csv",
    "Determinacao": DOWNLOAD_DIR / "Determinacao.csv",
    "Erros":        DOWNLOAD_DIR / "Erros.csv",
    "Pendentes":    DOWNLOAD_DIR / "Pendentes.csv",
}

OUT_XLSX = OUTPUT_DIR / "dashboard_unificado.xlsx"

# SGP: autodetect por padrão (altere se quiser fixar)
SGP_HEADER_ROW = None

# Engine Excel
try:
    import xlsxwriter  # noqa: F401
    ENGINE_WRITE = "xlsxwriter"
except Exception:
    ENGINE_WRITE = "openpyxl"
