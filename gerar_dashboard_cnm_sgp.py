#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Gera um dashboard de reconciliação entre CNM e SGP quando não há arquivos de eventos (SOA).
- Lê os arquivos CNM e SGP (Excel).
- Normaliza colunas e dados (tolerante a acentos, barras e pontuação).
- Evita KeyError para DATA/HORA usando colunas fallback.
- Cria STATUS: EM_AMBOS, SO_CNM, SO_SGP; TIPO = "BASE".
- Exporta CSV e HTML (DataTables) com filtro por Status.
"""

import os
import sys
import re
import unicodedata
import shutil
import pandas as pd
from pandas import to_datetime
from pathlib import Path
from textwrap import dedent

# =========================
# Configurações
# =========================
CNM_PATH = Path("download/Relatorio_CNM.xlsx")
SGP_PATH = Path("download/Relatorio_SGP.xlsx")
OUTPUT_DIR = Path("output")
MAX_ROWS_HEADER_SCAN = 25  # linhas para tentar detectar o header
MAX_ROWS_NORM = 2000       # linhas amostrais p/ descobrir header/normalizar preview

# =========================
# Utils de log
# =========================
def info(msg: str):
    print(f"[INFO] {msg}")

def warn(msg: str):
    print(f"[WARN] {msg}", file=sys.stderr)

def error(msg: str):
    print(f"[ERRO] {msg}", file=sys.stderr)

# =========================
# Normalização de string/coluna
# =========================
def strip_accents(s: str) -> str:
    return "".join(
        ch for ch in unicodedata.normalize("NFD", s)
        if unicodedata.category(ch) != "Mn"
    )

def norm_key(s: str) -> str:
    """
    Normaliza chave de comparação:
    - lower
    - remove acentos
    - troca qualquer coisa que não for [a-z0-9] por espaço
    - colapsa múltiplos espaços
    - remove espaços nas pontas
    Exemplos:
      'Nome/Razão Social' -> 'nome razao social'
      'cpf_cnpj'          -> 'cpf cnpj'
    """
    s = (s or "").strip().lower()
    s = strip_accents(s)
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

# =========================
# Normalização de células
# =========================
def _norm_cell(x):
    if pd.isna(x):
        return x
    if isinstance(x, str):
        s = x.strip()
        s = s.replace("\n", " ").replace("\r", " ").replace("\t", " ")
        while "  " in s:
            s = s.replace("  ", " ")
        return s
    return x

# =========================
# Leitura com detecção de header
# =========================
def detect_header_row(df_raw: pd.DataFrame, max_rows: int = MAX_ROWS_HEADER_SCAN) -> int:
    """
    Tenta detectar a linha do header procurando por uma linha que contenha
    indícios de nomes de colunas conhecidas (cpf, cpf_cnpj, nome, etc).
    Retorna índice 0-based da linha do header. Default: 0 se não achar.
    """
    keys = {"cpf", "cpf cnpj", "nome razao social", "cidade", "status", "tipo"}
    limit = min(len(df_raw), max_rows)
    for i in range(limit):
        row = df_raw.iloc[i].astype(str).str.lower().str.strip()
        norm_row_tokens = {norm_key(x) for x in row.tolist()}
        # Heurística: linha com muitos campos não-vazios e contendo alguma chave
        non_empty = (row != "") & (row != "nan")
        if non_empty.sum() >= 3:
            if any(k in " ".join(norm_row_tokens) for k in keys):
                return i
    # fallback comum nessas planilhas (observado em logs anteriores)
    return 8 if len(df_raw) > 9 else 0

def read_excel_with_header(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {path}")

    df_raw = pd.read_excel(path, header=None, engine="openpyxl")
    try:
        df_norm = df_raw.iloc[:MAX_ROWS_NORM].map(_norm_cell)  # pandas >=2.2
    except Exception:
        df_norm = df_raw.iloc[:MAX_ROWS_NORM].applymap(_norm_cell)

    header_row = detect_header_row(df_norm, MAX_ROWS_HEADER_SCAN)
    cols = df_raw.iloc[header_row].astype(str).str.strip().str.lower().tolist()
    df = df_raw.iloc[header_row+1:].copy()
    df.columns = cols

    info(f"header detectado na linha (0-based) {header_row}. Colunas: {df.columns.tolist()}")
    return df.reset_index(drop=True)

# =========================
# Helpers de coluna
# =========================
def pick_column(df: pd.DataFrame, candidates, label_for_error=None, required=True):
    """
    Tenta achar a primeira coluna existente em 'candidates', de forma TOLERANTE:
    - normaliza os nomes do DF e dos candidates (sem acento, sem pontuação, etc.)
    - aceita equivalências como 'nome/razão social' == 'nome_razao_social' == 'nome razao social'
    """
    # mapa normalizado -> nome original
    norm_map = {norm_key(c): c for c in df.columns}
    for cand in candidates:
        nk = norm_key(cand)
        if nk in norm_map:
            return norm_map[nk]

    if required:
        raise KeyError(f"[ERRO] Nao encontrei {label_for_error or candidates}. "
                       f"Candidates: {candidates}. Colunas disponiveis: {list(df.columns)}")
    return None

def clean_doc(val) -> str:
    """Remove qualquer caractere não numérico."""
    if pd.isna(val):
        return ""
    s = "".join(ch for ch in str(val) if ch.isdigit())
    return s

# =========================
# Carregar dados
# =========================
def carregar_dados():
    info("Lendo arquivos CNM e SGP...")
    cnm_df = read_excel_with_header(CNM_PATH)
    sgp_df = read_excel_with_header(SGP_PATH)

    # ----- Identifica colunas principais (tolerante a variações) -----
    col_doc_cnm = pick_column(
        cnm_df,
        ["cpf_cnpj", "cpf/cnpj", "cpf cnpj", "cpf"],
        "CPF/CNPJ (CNM)"
    )
    col_doc_sgp = pick_column(
        sgp_df,
        ["cpf_cnpj", "cpf/cnpj", "cpf cnpj", "cpf"],
        "CPF/CNPJ (SGP)"
    )

    col_nome_cnm = pick_column(
        cnm_df,
        ["nome_razao_social", "nome/razão social", "nome/razao social", "nome razao social", "nome", "razao_social"],
        "NOME (CNM)"
    )
    col_nome_sgp = pick_column(
        sgp_df,
        ["nome_razao_social", "nome/razão social", "nome/razao social", "nome razao social", "nome", "razao_social"],
        "NOME (SGP)"
    )

    # DATA/HORA com fallbacks (bases cadastrais não têm 'data_hora')
    col_data_cnm = pick_column(
        cnm_df,
        [
            "data_hora","data","data___hora","data__hora","data_e_hora",
            "data/_hora","data_/ _hora","data_/_hora",
            "data_cadastro_contrato","data cadastro contrato",
            "data_cadastro_cliente","data cadastro cliente",
            "vencimento",
        ],
        label_for_error="DATA/HORA (CNM)",
        required=False
    )
    col_data_sgp = pick_column(
        sgp_df,
        [
            "data_hora","data","data___hora","data__hora","data_e_hora",
            "data/_hora","data_/ _hora","data_/_hora",
            "data_cadastro_contrato","data cadastro contrato",
            "data_cadastro_cliente","data cadastro cliente",
            "vencimento",
        ],
        label_for_error="DATA/HORA (SGP)",
        required=False
    )

    # ----- Colunas normalizadas auxiliares -----
    cnm_df["_doc"] = cnm_df[col_doc_cnm].map(clean_doc)
    sgp_df["_doc"] = sgp_df[col_doc_sgp].map(clean_doc)

    cnm_df["_nome"] = cnm_df[col_nome_cnm].fillna("").astype(str).str.strip()
    sgp_df["_nome"] = sgp_df[col_nome_sgp].fillna("").astype(str).str.strip()

    if col_data_cnm:
        cnm_df["_data_hora"] = to_datetime(cnm_df[col_data_cnm], errors="coerce")
    else:
        cnm_df["_data_hora"] = pd.NaT

    if col_data_sgp:
        sgp_df["_data_hora"] = to_datetime(sgp_df[col_data_sgp], errors="coerce")
    else:
        sgp_df["_data_hora"] = pd.NaT

    return cnm_df, sgp_df

# =========================
# Montagem da tabela principal
# =========================
def montar_tabela_principal(cnm_df: pd.DataFrame, sgp_df: pd.DataFrame) -> pd.DataFrame:
    # Conjuntos de documentos
    cnm_docs = set(x for x in cnm_df["_doc"] if x)
    sgp_docs = set(x for x in sgp_df["_doc"] if x)

    em_ambos = cnm_docs & sgp_docs
    so_cnm   = cnm_docs - sgp_docs
    so_sgp   = sgp_docs - cnm_docs

    def classe_por_doc(doc: str) -> str:
        if doc in em_ambos: return "EM_AMBOS"
        if doc in so_cnm:   return "SO_CNM"
        if doc in so_sgp:   return "SO_SGP"
        return "DESCONHECIDO"

    def map_tipo_status(classe: str) -> tuple[str, str]:
        # Regras solicitadas para CNM_SGP:
        if classe == "EM_AMBOS":
            return ("NEGATIVADO", "NEGATIVADO")
        if classe in ("SO_SGP", "SO_CNM", "DESCONHECIDO"):
            return ("ERRO", "ERRO")
        return ("ERRO", "ERRO")

    def monta_linhas(df: pd.DataFrame, local_label: str) -> pd.DataFrame:
        out = pd.DataFrame()
        out["Documento"]   = df["_doc"]
        out["Nome"]        = df["_nome"]
        out["Data / Hora"] = df["_data_hora"]
        out["Local"]       = local_label

        # Classe (conciliacao)
        out["_classe"] = out["Documento"].map(classe_por_doc)

        # Aplica mapeamento Tipo/Status
        ts = out["_classe"].map(lambda c: map_tipo_status(c))
        out["Tipo"]   = ts.map(lambda x: x[0])
        out["Status"] = ts.map(lambda x: x[1])

        # Campo Operação vazio aqui (não há eventos)
        out["Operação"] = ""

        out = out[out["Documento"] != ""]
        return out[["Documento","Nome","Data / Hora","Operação","Tipo","Local","Status"]]

    t_cnm = monta_linhas(cnm_df, "CNM")
    t_sgp = monta_linhas(sgp_df, "SGP")

    tabela = pd.concat([t_cnm, t_sgp], ignore_index=True)
    tabela = tabela.drop_duplicates(subset=["Documento","Local"])
    tabela = tabela.sort_values(["Status","Documento","Local"]).reset_index(drop=True)
    tabela.insert(0, "Id", range(1, len(tabela) + 1))
    return tabela

# =========================
# Export HTML (DataTables)
# =========================
def escrever_html_datatable(df: pd.DataFrame, path_html: Path, titulo: str = "Dashboard CNM x SGP"):
    info(f"Criando HTML: {path_html}")
    df_html = df.copy()
    if "Data / Hora" in df_html.columns:
        df_html["Data / Hora"] = df_html["Data / Hora"].astype(str).replace("NaT", "")

    html = f"""\
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8"/>
<title>{titulo}</title>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<link rel="stylesheet" href="https://cdn.datatables.net/1.13.8/css/jquery.dataTables.min.css"/>
<style>
  body {{ font-family: system-ui, -apple-system, 'Segoe UI', Roboto, Arial, sans-serif; padding: 18px; }}
  h1 {{ margin-top: 0; }}
  .filters {{ margin-bottom: 12px; display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }}
  .subtitle {{ margin: 6px 0 16px; color: #444; }}
  table.dataTable thead th, table.dataTable tfoot th {{ font-weight: 600; }}
</style>
</head>
<body>
  <h1>{titulo}</h1>
  <div class="filters">
    <label for="statusFilter"><strong>Status:</strong></label>
    <select id="statusFilter">
      <option value="">(Todos)</option>
      <option>EM_AMBOS</option>
      <option>SO_CNM</option>
      <option>SO_SGP</option>
      <option>DESCONHECIDO</option>
    </select>
  </div>
  <div class="subtitle" id="subtitle"></div>

  <table id="tbl" class="display" style="width:100%">
    <thead>
      <tr>{"".join(f"<th>{c}</th>" for c in df_html.columns)}</tr>
    </thead>
    <tbody>
      {"".join("<tr>" + "".join(f"<td>{(str(r[c]) if pd.notna(r[c]) else '')}</td>" for c in df_html.columns) + "</tr>" for _, r in df_html.iterrows())}
    </tbody>
  </table>

  <script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
  <script src="https://cdn.datatables.net/1.13.8/js/jquery.dataTables.min.js"></script>
  <script>
    const SUBTITULOS = {{
      "": "Mostrando todos os registros conciliados entre CNM e SGP.",
      "EM_AMBOS": "Registros presentes em ambas as bases (CNM e SGP).",
      "SO_CNM": "Registros presentes apenas na base CNM.",
      "SO_SGP": "Registros presentes apenas na base SGP.",
      "DESCONHECIDO": "Registros cuja presença não pôde ser determinada."
    }};

    function atualizarSubtitulo(valor) {{
      const el = document.getElementById('subtitle');
      el.textContent = SUBTITULOS[valor ?? ""] || SUBTITULOS[""];
    }}

    $(document).ready(function() {{
      const table = $('#tbl').DataTable({{
        pageLength: 25,
        order: [[1, 'asc']],
        language: {{
          url: 'https://cdn.datatables.net/plug-ins/1.13.8/i18n/pt-BR.json'
        }}
      }});

      atualizarSubtitulo("");

      $('#statusFilter').on('change', function() {{
        const val = this.value;
        const statusIdx = $('#tbl thead th').toArray().findIndex(th => th.textContent.trim() === 'Status');
        if (statusIdx >= 0) {{
          table.column(statusIdx).search(val ? '^' + val + '$' : '', true, false).draw();
        }}
        atualizarSubtitulo(val);
      }});
    }});
  </script>
</body>
</html>
"""
    path_html.write_text(dedent(html), encoding="utf-8")

# =========================
# Main
# =========================
def limpar_output():
    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def gerar_dashboard():
    try:
        limpar_output()
        cnm_df, sgp_df = carregar_dados()
        tabela = montar_tabela_principal(cnm_df, sgp_df)

        csv_path = OUTPUT_DIR / "tabela_principal.csv"
        tabela.to_csv(csv_path, index=False, encoding="utf-8")
        info(f"CSV gerado: {csv_path}")

        html_path = OUTPUT_DIR / "dashboard_cnm_sgp.html"
        escrever_html_datatable(tabela, html_path, "Dashboard CNM × SGP (Reconciliação de Base)")
        info("Concluído.")
    except Exception as e:
        error(str(e))
        raise

if __name__ == "__main__":
    gerar_dashboard()
